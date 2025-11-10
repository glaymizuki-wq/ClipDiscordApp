using System;
using System.IO;
using System.Text.Json;
using ClipDiscordApp.Models;

namespace ClipDiscordApp.Utils
{
    public static class SettingsStore
    {
        private static readonly string Dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ClipDiscordApp");
        private static readonly string FilePath = Path.Combine(Dir, "settings.json");

        public static AppSettings Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return null;
                var json = File.ReadAllText(FilePath);
                return JsonSerializer.Deserialize<AppSettings>(json);
            }
            catch { return null; }
        }

        public static void Save(AppSettings settings)
        {
            try
            {
                if (!Directory.Exists(Dir)) Directory.CreateDirectory(Dir);
                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(FilePath, json);
            }
            catch { }
        }
    }
}