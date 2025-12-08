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

                // 日本語OCR用に基本的な正規化
                var s = raw.Trim();
                // 全角スペースを半角に
                s = s.Replace('　', ' ');
                // よくあるOCR誤認補正（例: 洛→落、昇→上昇）
                s = s.Replace("洛", "落").Replace("昇", "上昇");

                // トークン分割（スペースや記号で分割）
                var tokens = s.Split(new[] { ' ', ':', '-', 'T' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var cand in tokens)
                {
                    if (string.IsNullOrWhiteSpace(cand)) continue;

                    // 「下落中」「上昇中」とのスコアを計算
                    var scoreDown = ComputeLabelScore(cand, "下落中");
                    var scoreUp = ComputeLabelScore(cand, "上昇中");

                    results.Add(new LabelCandidate(cand, "下落中", scoreDown, source));
                    results.Add(new LabelCandidate(cand, "上昇中", scoreUp, source));

                    System.Diagnostics.Debug.WriteLine(
                        $"[OcrHelpers] Candidate src={source} token='{cand}' 下落中={scoreDown:F3} 上昇中={scoreUp:F3}"
                    );
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