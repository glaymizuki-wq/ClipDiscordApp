using ClipDiscordApp.Models; // OcrResult, OcrWord, ExtractRule 等
using System.Text.RegularExpressions;

public class ExtractResult
{
    public string RuleId { get; set; } = "";
    public string RuleName { get; set; } = "";
    public List<string> Matches { get; set; } = new();
}

public static class OcrParser
{
    private const int FuzzyThreshold = 1;

    private static string NormalizeForComparison(string s)
    {
        if (string.IsNullOrEmpty(s)) return string.Empty;
        var norm = s.Normalize(System.Text.NormalizationForm.FormKC);
        norm = System.Text.RegularExpressions.Regex.Replace(norm, @"\s+", " ").Trim();
        return norm.ToUpperInvariant();
    }

    private static string NormalizeForOutput(string s)
    {
        if (string.IsNullOrEmpty(s)) return string.Empty;
        return s.Normalize(System.Text.NormalizationForm.FormKC).Trim();
    }

    private static int LevenshteinDistance(string a, string b)
    {
        if (a == null) a = "";
        if (b == null) b = "";
        var n = a.Length;
        var m = b.Length;
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

    // 新規追加: 文字列オーバーロード（読み取り専用 RawText の代わりにこれを使う）
    public static List<ExtractMatch> ParseByRules(string rawText, IEnumerable<ExtractRule> rules)
    {
        var results = new List<ExtractMatch>();

        var compText = NormalizeForComparison(rawText ?? string.Empty);

        System.Diagnostics.Debug.WriteLine($"[ParseByRules] OCR Raw length={(rawText?.Length ?? 0)}");
        System.Diagnostics.Debug.WriteLine($"[ParseByRules] OCR Raw: '{rawText}'");
        System.Diagnostics.Debug.WriteLine($"[ParseByRules] OCR Norm: '{compText}'");

        foreach (var rule in (rules ?? Enumerable.Empty<ExtractRule>()).Where(r => r.Enabled).OrderBy(r => r.Order))
        {
            var em = new ExtractMatch { RuleId = rule.Id, RuleName = rule.Name };
            var rawPattern = rule.Pattern ?? string.Empty;
            var compPattern = NormalizeForComparison(rawPattern);

            if (rule.Type == ExtractRuleType.Keyword)
            {
                try
                {
                    int idx = compText.IndexOf(compPattern, StringComparison.Ordinal);
                    if (idx >= 0)
                    {
                        // なるべく原文に近い部分を格納するため rawText 側でも検索
                        var rawUpper = NormalizeForComparison(rawText ?? string.Empty);
                        int rawIdx = rawUpper.IndexOf(compPattern, StringComparison.Ordinal);
                        if (rawIdx >= 0 && rawIdx + compPattern.Length <= (rawText ?? "").Length)
                        {
                            var matchedOriginal = (rawText ?? "").Substring(rawIdx, compPattern.Length);
                            em.Matches.Add(NormalizeForOutput(matchedOriginal));
                        }
                        else
                        {
                            em.Matches.Add(NormalizeForOutput(rawPattern));
                        }

                        System.Diagnostics.Debug.WriteLine($"[ParseByRules] Keyword matched rule='{rule.Name}' pattern='{rawPattern}'");
                        results.Add(em);
                        continue;
                    }

                    var tokens = compPattern.Split(' ').Where(t => !string.IsNullOrEmpty(t)).ToArray();
                    if (tokens.Length > 1)
                    {
                        int pos = 0;
                        bool allFound = true;
                        foreach (var t in tokens)
                        {
                            int p = compText.IndexOf(t, pos, StringComparison.Ordinal);
                            if (p < 0) { allFound = false; break; }
                            pos = p + t.Length;
                        }
                        if (allFound)
                        {
                            em.Matches.Add(NormalizeForOutput(rawPattern));
                            results.Add(em);
                            System.Diagnostics.Debug.WriteLine($"[ParseByRules] Keyword (token-match) rule='{rule.Name}' pattern='{rawPattern}'");
                            continue;
                        }
                    }

                    if (!string.IsNullOrEmpty(compPattern))
                    {
                        var pLen = compPattern.Length;
                        var textLen = compText.Length;
                        int bestDist = int.MaxValue;
                        int bestPos = -1;
                        int maxWindow = Math.Min(2000, Math.Max(textLen, pLen + 20));
                        for (int i = 0; i < Math.Max(1, textLen - Math.Max(0, pLen - 1)); i++)
                        {
                            int len = Math.Min(pLen + FuzzyThreshold, textLen - i);
                            if (len <= 0) break;
                            string window = compText.Substring(i, len);
                            int dist = LevenshteinDistance(window, compPattern);
                            if (dist < bestDist)
                            {
                                bestDist = dist;
                                bestPos = i;
                                if (bestDist == 0) break;
                            }
                            if (i > maxWindow) break;
                        }

                        if (bestDist <= FuzzyThreshold)
                        {
                            var rawUpper = NormalizeForComparison(rawText ?? string.Empty);
                            int rawIdx = Math.Max(0, Math.Min(rawUpper.Length - compPattern.Length, bestPos));
                            if (rawIdx >= 0 && rawIdx + compPattern.Length <= (rawText ?? "").Length)
                            {
                                var matchedOriginal = (rawText ?? "").Substring(rawIdx, Math.Min(compPattern.Length, (rawText ?? "").Length - rawIdx));
                                em.Matches.Add(NormalizeForOutput(matchedOriginal));
                            }
                            else
                            {
                                em.Matches.Add(NormalizeForOutput(rawPattern));
                            }

                            System.Diagnostics.Debug.WriteLine($"[ParseByRules] Keyword fuzzy matched rule='{rule.Name}' pattern='{rawPattern}' dist={bestDist}");
                            results.Add(em);
                            continue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ParseByRules] Keyword match error rule='{rule.Name}': {ex}");
                }
            }
            else
            {
                try
                {
                    var pattern = rawPattern ?? string.Empty;
                    var rx = new System.Text.RegularExpressions.Regex(pattern,
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase |
                        System.Text.RegularExpressions.RegexOptions.CultureInvariant);

                    foreach (System.Text.RegularExpressions.Match m in rx.Matches(compText))
                    {
                        if (m.Success)
                        {
                            var matchedComp = m.Value;
                            var rawUpper = NormalizeForComparison(rawText ?? string.Empty);
                            int rawIdx = rawUpper.IndexOf(matchedComp, StringComparison.Ordinal);
                            if (rawIdx >= 0 && rawIdx + matchedComp.Length <= (rawText ?? "").Length)
                            {
                                em.Matches.Add(NormalizeForOutput((rawText ?? "").Substring(rawIdx, matchedComp.Length)));
                            }
                            else
                            {
                                em.Matches.Add(NormalizeForOutput(matchedComp));
                            }
                        }
                    }

                    if (em.Matches.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ParseByRules] Regex matched rule='{rule.Name}' pattern='{pattern}' matches={em.Matches.Count}");
                        results.Add(em);
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ParseByRules] Invalid regex for rule '{rule.Name}': {ex.Message}");
                }
            }

            System.Diagnostics.Debug.WriteLine($"[ParseByRules] No match for rule='{rule.Name}' pattern='{rawPattern}'");
        }

        System.Diagnostics.Debug.WriteLine($"[ParseByRules] Finished. totalMatches={results.Count}");
        return results;
    }

    // 既存の OcrResult オーバーロードは内部で文字列版を呼ぶようにする
    public static List<ExtractMatch> ParseByRules(OcrResult ocr, IEnumerable<ExtractRule> rules)
    {
        // ocr?.RawText は読み取り専用でも OK（参照するだけ）
        var raw = ocr?.RawText ?? string.Empty;
        return ParseByRules(raw, rules);
    }
}