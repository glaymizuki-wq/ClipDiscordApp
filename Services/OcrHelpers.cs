using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using ClipDiscordApp.Models;
using Tesseract;

namespace ClipDiscordApp.Services
{
    public static class OcrHelpers
    {
        // PSM 順序: SingleLine を優先
        public static string DoOcrWithRetries(TesseractEngine engine, Pix pix)
        {
            if (engine == null || pix == null) return string.Empty;

            // PageSegMode の順序を SingleLine 優先にしている
            var psmCandidates = new[] { PageSegMode.SingleLine, PageSegMode.SingleWord, PageSegMode.Auto, PageSegMode.SingleBlock };
            string best = string.Empty;

            foreach (var psm in psmCandidates)
            {
                try
                {
                    engine.DefaultPageSegMode = psm;
                    using var page = engine.Process(pix);
                    var text = page.GetText()?.Trim() ?? string.Empty;
                    System.Diagnostics.Debug.WriteLine($"[OcrHelpers] PSM={psm} -> '{text}' (len={text.Length})");
                    try { System.Diagnostics.Debug.WriteLine($"[OcrHelpers] MeanConfidence={page.GetMeanConfidence():F2}"); } catch { }
                    // 最初に見つかった non-empty を優先返却（必要ならスコアリング選択に拡張）
                    if (!string.IsNullOrWhiteSpace(text) && string.IsNullOrWhiteSpace(best)) best = text;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[OcrHelpers] PSM {psm} exception: {ex.Message}");
                }
            }
            return best;
        }

        // raws からトークン化して BUY/SELL を評価し候補リストで返す
        public static List<LabelCandidate> NormalizeAndDetectLabels(IEnumerable<string> raws, string source)
        {
            var results = new List<LabelCandidate>();
            if (raws == null) return results;

            foreach (var raw in raws)
            {
                if (string.IsNullOrWhiteSpace(raw)) continue;

                var s = raw.Trim().ToUpperInvariant();
                s = s.Replace('—', '-').Replace('ー', '-').Replace('　', ' ');
                s = new string(s.Where(c => char.IsLetterOrDigit(c) || c == ':' || c == '-' || char.IsWhiteSpace(c)).ToArray());

                var tokens = s.Split(new[] { ':', '-', ' ', 'T' }, StringSplitOptions.RemoveEmptyEntries);

                var tokenCandidates = new HashSet<string>();
                foreach (var t in tokens)
                {
                    tokenCandidates.Add(t);
                    tokenCandidates.Add(t.Replace('O', '0').Replace('I', '1').Replace('L', '1'));
                    tokenCandidates.Add(t.Replace('0', 'O').Replace('1', 'I'));
                    tokenCandidates.Add(t.Replace('S', '5').Replace('B', '8'));
                    tokenCandidates.Add(t.Replace('5', 'S').Replace('8', 'B'));
                    if (t.All(c => c == 'O' || c == '0')) tokenCandidates.Add(new string(t.Select(c => c == 'O' ? '0' : 'O').ToArray()));
                }

                foreach (var cand in tokenCandidates)
                {
                    if (string.IsNullOrWhiteSpace(cand)) continue;
                    var scoreSell = ComputeLabelScore(cand, "SELL");
                    var scoreBuy = ComputeLabelScore(cand, "BUY");
                    var sellC = new LabelCandidate(cand, "SELL", scoreSell, source);
                    var buyC = new LabelCandidate(cand, "BUY", scoreBuy, source);
                    results.Add(sellC);
                    results.Add(buyC);
                    System.Diagnostics.Debug.WriteLine($"[OcrHelpers] Candidate src={source} token='{cand}' SELL={scoreSell:F3} BUY={scoreBuy:F3}");
                }
            }

            return results;
        }

        // 0..1 の簡易スコア
        private static double ComputeLabelScore(string token, string label)
        {
            if (string.IsNullOrWhiteSpace(token) || string.IsNullOrWhiteSpace(label)) return 0.0;
            var tok = token.ToUpperInvariant();
            var whitelist = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var whitelistCount = tok.Count(c => whitelist.Contains(c));
            var baseScore = whitelistCount * 100 + tok.Length * 10;

            var dist = Levenshtein(tok, label);
            var maxLen = Math.Max(tok.Length, label.Length);
            var fuzzy = maxLen == 0 ? 0.0 : 1.0 - (double)dist / maxLen; // 1.0 完全一致

            var scaledBase = Math.Min(1.0, baseScore / 500.0);
            var combined = scaledBase * 0.5 + fuzzy * 0.5;
            return Math.Max(0.0, Math.Min(1.0, combined));
        }

        // Levenshtein
        private static int Levenshtein(string a, string b)
        {
            if (a == null) return b?.Length ?? 0;
            if (b == null) return a.Length;
            var n = a.Length;
            var m = b.Length;
            var d = new int[n + 1, m + 1];
            for (int i = 0; i <= n; i++) d[i, 0] = i;
            for (int j = 0; j <= m; j++) d[0, j] = j;
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    var cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }
            return d[n, m];
        }
    }
}