using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using ClipDiscordApp.Models;

namespace ClipDiscordApp.Services
{
    public static class Notifier
    {
        // HttpClient は使い回す
        private static readonly HttpClient _discordHttpClient = new HttpClient();

        // WebhookUrl は外部ファイルから取得して設定する
        public static string WebhookUrl { get; private set; } = string.Empty;

        // 初期化処理で呼び出す
        public static void InitializeWebhook(string key = "ClipDiscordApp")
        {
            WebhookUrl = GetWebhookUrl(key);
        }

        // MatchResultItem を受け取って送信
        public static async Task<bool> SendOrderAsync(MatchResultItem item)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(WebhookUrl))
                {
                    System.Diagnostics.Debug.WriteLine("[Notifier] WebhookUrl not set");
                    return false;
                }

                var text = (item.Matches != null && item.Matches.Length > 0)
                    ? string.Join(",", item.Matches)
                    : item.RuleName?.ToUpperInvariant();

                // 変更後（小文字や余分な空白を避けたいなら ToUpperInvariant を使う）
                var payload = new { content = (text ?? item.RuleName ?? string.Empty).ToUpperInvariant() }; var json = JsonSerializer.Serialize(payload);
                using var body = new StringContent(json, Encoding.UTF8, "application/json");
                var resp = await _discordHttpClient.PostAsync(WebhookUrl, body);
                var respText = await resp.Content.ReadAsStringAsync();

                System.Diagnostics.Debug.WriteLine($"[Notifier] status={(int)resp.StatusCode} body={respText}");

                return resp.IsSuccessStatusCode;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Notifier] SendOrderAsync failed: {ex}");
                return false;
            }
        }

        // 既存の GetWebhookUrl を利用
        private static string GetWebhookUrl(string key = "ClipDiscordApp")
        {
            try
            {
                if (string.IsNullOrWhiteSpace(key)) key = "ClipDiscordApp";

                var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                var appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ClipDiscordApp");

                var candidates = new[]
                {
                    Path.Combine(exeDir, "discord_webhooks.json"),
                    Path.Combine(appDataDir, "discord_webhooks.json")
                };

                foreach (var configPath in candidates)
                {
                    if (!File.Exists(configPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[GetWebhookUrl] config not found: {configPath}");
                        continue;
                    }

                    string json;
                    try
                    {
                        json = File.ReadAllText(configPath, Encoding.UTF8);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[GetWebhookUrl] could not read {configPath}: {ex}");
                        continue;
                    }

                    try
                    {
                        using var doc = JsonDocument.Parse(json);
                        if (doc.RootElement.TryGetProperty(key, out var val) && val.ValueKind == JsonValueKind.String)
                        {
                            var url = val.GetString()?.Trim();
                            if (!string.IsNullOrEmpty(url) && Uri.TryCreate(url, UriKind.Absolute, out var u) && (u.Scheme == "https" || u.Scheme == "http"))
                            {
                                System.Diagnostics.Debug.WriteLine($"[GetWebhookUrl] loaded webhook from {configPath}");
                                return url;
                            }
                            System.Diagnostics.Debug.WriteLine($"[GetWebhookUrl] invalid or empty url for key '{key}' in {configPath}");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[GetWebhookUrl] key '{key}' not found or not a string in {configPath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[GetWebhookUrl] parse error in {configPath}: {ex}");
                    }
                }

                // 環境変数フォールバック
                var envKey = $"DISCORD_WEBHOOK_{key.ToUpperInvariant()}";
                var envUrl = Environment.GetEnvironmentVariable(envKey);
                if (!string.IsNullOrWhiteSpace(envUrl) && Uri.TryCreate(envUrl.Trim(), UriKind.Absolute, out var envUri) && (envUri.Scheme == "https" || envUri.Scheme == "http"))
                {
                    System.Diagnostics.Debug.WriteLine($"[GetWebhookUrl] loaded webhook from ENV {envKey}");
                    return envUrl.Trim();
                }

                var fallbackEnv = Environment.GetEnvironmentVariable("DISCORD_WEBHOOK_CLIPDISCORDAPP");
                if (!string.IsNullOrWhiteSpace(fallbackEnv) && Uri.TryCreate(fallbackEnv.Trim(), UriKind.Absolute, out var fbUri) && (fbUri.Scheme == "https" || fbUri.Scheme == "http"))
                {
                    System.Diagnostics.Debug.WriteLine("[GetWebhookUrl] loaded webhook from ENV DISCORD_WEBHOOK_CLIPDISCORDAPP");
                    return fallbackEnv.Trim();
                }

                System.Diagnostics.Debug.WriteLine("[GetWebhookUrl] webhook not found in any candidate locations");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetWebhookUrl] unexpected error: {ex}");
            }

            return string.Empty;
        }
    }
}