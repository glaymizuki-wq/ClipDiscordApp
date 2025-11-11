using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClipDiscordApp.Utils
{
    public class OcrTextUtils
    {
        // 既存メソッド群...

        // 統一された閾値プロパティ（実行時に変更可能）
        public static int FuzzyThreshold { get; set; } = 2;

        // ラッパー（将来ロジックを追加したい場合はここに追加）
        public static int GetFuzzyThreshold() => FuzzyThreshold;

        public static string NormalizeForComparison(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = s.ToUpperInvariant();

            // 基本の不要文字除去（コロンやハイフンは残す）
            var sb = new System.Text.StringBuilder();
            foreach (var ch in s)
            {
                if (char.IsLetterOrDigit(ch) || ch == ':' || ch == '-' || ch == ' ') sb.Append(ch);
            }
            s = sb.ToString();

            // よくある置換（順序は重要）
            s = s.Replace('0', 'O');   // 0 -> O
            s = s.Replace('1', 'I');   // 1 -> I
            s = s.Replace('5', 'S');   // 5 -> S
            s = s.Replace('8', 'B');   // 8 -> B
            s = s.Replace('|', 'I');   // | -> I
            s = s.Replace("::", ":");  // 重複コロン修正
            s = s.Replace(" ", "");    // 比較用は空白除去

            // ---------- TEST ONLY: common OCR misrecognitions mapping ----------
            // テストが終わったらこのブロックは必ず削除してください
            s = s.Replace("SML", "SELL");
            s = s.Replace("SMI", "SELL");
            s = s.Replace("SEI", "SELL");
            s = s.Replace("S M L", "SELL");
            s = s.Replace("0SE", "OSE");
            return s;
        }

        public static string NormalizeForOutput(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            // 最小限のトリムで表示向けに
            return s.Trim();
        }

        public static void LogFuzzyMatch(string token, string pattern, int dist)
        {
            System.Diagnostics.Debug.WriteLine($"[Fuzzy] token='{token}' pattern='{pattern}' dist={dist}");
        }

        public static int LevenshteinDistance(string a, string b)
        {
            if (a == null) a = "";
            if (b == null) b = "";
            int n = a.Length, m = b.Length;
            if (n == 0) return m;
            if (m == 0) return n;
            var d = new int[n + 1, m + 1];
            for (int i = 0; i <= n; i++) d[i, 0] = i;
            for (int j = 0; j <= m; j++) d[0, j] = j;
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }
            return d[n, m];
        }
    }
}