using ClipDiscordApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClipDiscordApp.Parsers
{
    public static class OcrParser
    {
        public static List<ExtractMatch> ParseByRules(string rawText, IEnumerable<ExtractRule> rules)
        {
            var results = new List<ExtractMatch>();

            // --- last-token first check (日時 + ラベル 形式向け) ---
            if (!string.IsNullOrWhiteSpace(rawText))
            {
                var rawTokens = rawText.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                var lastRaw = rawTokens.LastOrDefault();
                if (!string.IsNullOrWhiteSpace(lastRaw))
                {
                    // 比較用正規化（NormalizeForComparison が日本語を崩す場合は lastRaw 直接比較に切り替える）
                    var cand = ClipDiscordApp.Utils.OcrTextUtils.NormalizeForComparison(lastRaw);

                    // --- Japanese labels ---
                    if (cand == "上昇中")
                    {
                        var upRule = (rules ?? Enumerable.Empty<ExtractRule>())
                            .FirstOrDefault(r => r.Enabled && r.Type == ExtractRuleType.Keyword && r.Pattern == "上昇中");
                        var em = upRule != null
                            ? new ExtractMatch { RuleId = upRule.Id, RuleName = upRule.Name }
                            : new ExtractMatch { RuleId = "up", RuleName = "上昇中 (auto)" };
                        em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(lastRaw));
                        results.Add(em);
                        System.Diagnostics.Debug.WriteLine($"[ParseByRules] LastToken matched '上昇中' lastRaw='{lastRaw}'");
                        return results;
                    }

                    if (cand == "下落中")
                    {
                        var downRule = (rules ?? Enumerable.Empty<ExtractRule>())
                            .FirstOrDefault(r => r.Enabled && r.Type == ExtractRuleType.Keyword && r.Pattern == "下落中");
                        var em = downRule != null
                            ? new ExtractMatch { RuleId = downRule.Id, RuleName = downRule.Name }
                            : new ExtractMatch { RuleId = "down", RuleName = "下落中 (auto)" };
                        em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(lastRaw));
                        results.Add(em);
                        System.Diagnostics.Debug.WriteLine($"[ParseByRules] LastToken matched '下落中' lastRaw='{lastRaw}'");
                        return results;
                    }

                    // --- English BUY/SELL legacy checks (後方互換のため残置) ---
                    // quick regex checks
                    if (System.Text.RegularExpressions.Regex.IsMatch(cand, @"\bS[EI1\|]?[L1I]{1,2}\b"))
                    {
                        var sellRule = (rules ?? Enumerable.Empty<ExtractRule>()).FirstOrDefault(r => r.Enabled && r.Type == ExtractRuleType.Keyword && string.Equals(r.Pattern, "SELL", StringComparison.OrdinalIgnoreCase));
                        var em = sellRule != null ? new ExtractMatch { RuleId = sellRule.Id, RuleName = sellRule.Name } : new ExtractMatch { RuleId = "sell", RuleName = "Sell (auto)" };
                        em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(lastRaw));
                        results.Add(em);
                        System.Diagnostics.Debug.WriteLine($"[ParseByRules] LastToken regex matched SELL cand='{cand}' lastRaw='{lastRaw}'");
                        return results;
                    }

                    if (System.Text.RegularExpressions.Regex.IsMatch(cand, @"\bB[0O]?[UY1I]?\b"))
                    {
                        var buyRule = (rules ?? Enumerable.Empty<ExtractRule>()).FirstOrDefault(r => r.Enabled && r.Type == ExtractRuleType.Keyword && string.Equals(r.Pattern, "BUY", StringComparison.OrdinalIgnoreCase));
                        var em = buyRule != null ? new ExtractMatch { RuleId = buyRule.Id, RuleName = buyRule.Name } : new ExtractMatch { RuleId = "buy", RuleName = "Buy (auto)" };
                        em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(lastRaw));
                        results.Add(em);
                        System.Diagnostics.Debug.WriteLine($"[ParseByRules] LastToken regex matched BUY cand='{cand}' lastRaw='{lastRaw}'");
                        return results;
                    }

                    // fuzzy checks (Levenshtein)
                    var distSell = ClipDiscordApp.Utils.OcrTextUtils.LevenshteinDistance(cand, "SELL");
                    var distBuy = ClipDiscordApp.Utils.OcrTextUtils.LevenshteinDistance(cand, "BUY");
                    ClipDiscordApp.Utils.OcrTextUtils.LogFuzzyMatch(cand, "SELL", distSell);
                    ClipDiscordApp.Utils.OcrTextUtils.LogFuzzyMatch(cand, "BUY", distBuy);

                    int fuzzy = ClipDiscordApp.Utils.OcrTextUtils.GetFuzzyThreshold();
                    if (distSell <= Math.Max(1, fuzzy))
                    {
                        var sellRule = (rules ?? Enumerable.Empty<ExtractRule>()).FirstOrDefault(r => r.Enabled && r.Type == ExtractRuleType.Keyword && string.Equals(r.Pattern, "SELL", StringComparison.OrdinalIgnoreCase));
                        var em = sellRule != null ? new ExtractMatch { RuleId = sellRule.Id, RuleName = sellRule.Name } : new ExtractMatch { RuleId = "sell", RuleName = "Sell (fuzzy)" };
                        em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(lastRaw));
                        results.Add(em);
                        System.Diagnostics.Debug.WriteLine($"[ParseByRules] LastToken fuzzy matched SELL cand='{cand}' dist={distSell}");
                        return results;
                    }

                    if (distBuy <= Math.Max(1, fuzzy - 1))
                    {
                        var buyRule = (rules ?? Enumerable.Empty<ExtractRule>()).FirstOrDefault(r => r.Enabled && r.Type == ExtractRuleType.Keyword && string.Equals(r.Pattern, "BUY", StringComparison.OrdinalIgnoreCase));
                        var em = buyRule != null ? new ExtractMatch { RuleId = buyRule.Id, RuleName = buyRule.Name } : new ExtractMatch { RuleId = "buy", RuleName = "Buy (fuzzy)" };
                        em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(lastRaw));
                        results.Add(em);
                        System.Diagnostics.Debug.WriteLine($"[ParseByRules] LastToken fuzzy matched BUY cand='{cand}' dist={distBuy}");
                        return results;
                    }

                    // 保険: 末尾 L が欠けている可能性（例: "SML" -> "SELL"）
                    if (distSell > Math.Max(1, fuzzy))
                    {
                        var cand2 = cand.Replace(" ", ""); // "S M L" などを詰める
                        if (cand2.Length <= 4 && (cand2.EndsWith("SM") || cand2.EndsWith("S") || cand2.EndsWith("SE") || cand2.EndsWith("SML")))
                        {
                            var withL = cand2 + "L";
                            var dist2 = ClipDiscordApp.Utils.OcrTextUtils.LevenshteinDistance(withL, "SELL");
                            ClipDiscordApp.Utils.OcrTextUtils.LogFuzzyMatch(withL, "SELL", dist2);
                            if (dist2 <= Math.Max(1, fuzzy))
                            {
                                var sellRule = (rules ?? Enumerable.Empty<ExtractRule>()).FirstOrDefault(r => r.Enabled && r.Type == ExtractRuleType.Keyword && string.Equals(r.Pattern, "SELL", StringComparison.OrdinalIgnoreCase));
                                var em = sellRule != null ? new ExtractMatch { RuleId = sellRule.Id, RuleName = sellRule.Name } : new ExtractMatch { RuleId = "sell", RuleName = "Sell (auto-lfix)" };
                                em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(lastRaw));
                                results.Add(em);
                                System.Diagnostics.Debug.WriteLine($"[ParseByRules] LastToken fixed-by-appendL matched SELL cand='{cand}' lastRaw='{lastRaw}'");
                                return results;
                            }
                        }
                    }
                }
            }
            // --- end last-token first check ---

            // 比較用正規化文字列
            var compText = ClipDiscordApp.Utils.OcrTextUtils.NormalizeForComparison(rawText ?? string.Empty);

            System.Diagnostics.Debug.WriteLine($"[ParseByRules] OCR Raw length={(rawText?.Length ?? 0)}");
            System.Diagnostics.Debug.WriteLine($"[ParseByRules] OCR Raw: '{rawText}'");
            System.Diagnostics.Debug.WriteLine($"[ParseByRules] OCR Norm: '{compText}'");

            // トークン化（: - 空白で分割）※重複を除く
            var compTokens = compText
                .Split(new[] { ':', '-', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(t => !string.IsNullOrWhiteSpace(t))
                .Distinct()
                .ToArray();

            // 統一された閾値取得
            int fuzzyThreshold = ClipDiscordApp.Utils.OcrTextUtils.GetFuzzyThreshold();

            foreach (var rule in (rules ?? Enumerable.Empty<ExtractRule>()).Where(r => r.Enabled).OrderBy(r => r.Order))
            {
                var em = new ExtractMatch { RuleId = rule.Id, RuleName = rule.Name };
                var rawPattern = rule.Pattern ?? string.Empty;
                var compPattern = ClipDiscordApp.Utils.OcrTextUtils.NormalizeForComparison(rawPattern);

                if (rule.Type == ExtractRuleType.Keyword)
                {
                    try
                    {
                        // 1) 完全一致（比較用文字列）
                        int idx = compText.IndexOf(compPattern, StringComparison.Ordinal);
                        if (idx >= 0)
                        {
                            var rawUpper = ClipDiscordApp.Utils.OcrTextUtils.NormalizeForComparison(rawText ?? string.Empty);
                            int rawIdx = rawUpper.IndexOf(compPattern, StringComparison.Ordinal);
                            if (rawIdx >= 0 && rawIdx + compPattern.Length <= (rawText ?? "").Length)
                            {
                                var matchedOriginal = (rawText ?? "").Substring(rawIdx, compPattern.Length);
                                em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(matchedOriginal));
                            }
                            else
                            {
                                em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(rawPattern));
                            }

                            System.Diagnostics.Debug.WriteLine($"[ParseByRules] Keyword matched rule='{rule.Name}' pattern='{rawPattern}'");
                            results.Add(em);
                            continue;
                        }

                        // 2) トークン単位の厳密一致
                        if (compTokens.Any(t => t == compPattern))
                        {
                            em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(rawPattern));
                            results.Add(em);
                            System.Diagnostics.Debug.WriteLine($"[ParseByRules] Token exact matched rule='{rule.Name}' pattern='{rawPattern}'");
                            continue;
                        }

                        // 3) トークン単位ファジーマッチ（Levenshtein）
                        foreach (var token in compTokens)
                        {
                            var dist = ClipDiscordApp.Utils.OcrTextUtils.LevenshteinDistance(token, compPattern);
                            ClipDiscordApp.Utils.OcrTextUtils.LogFuzzyMatch(token, compPattern, dist);
                            var threshold = Math.Max(1, compPattern.Length <= 4 ? 1 : fuzzyThreshold);
                            if (dist <= threshold)
                            {
                                em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(rawPattern));
                                results.Add(em);
                                System.Diagnostics.Debug.WriteLine($"[ParseByRules] Token fuzzy matched rule='{rule.Name}' pattern='{rawPattern}' token='{token}' dist={dist}");
                                goto NextRule;
                            }
                        }

                        // 4) 複数語のトークン順序マッチ
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
                                em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(rawPattern));
                                results.Add(em);
                                System.Diagnostics.Debug.WriteLine($"[ParseByRules] Keyword (token-match) rule='{rule.Name}' pattern='{rawPattern}'");
                                continue;
                            }
                        }

                        // 5) ウィンドウ走査によるファジーマッチ
                        if (!string.IsNullOrEmpty(compPattern))
                        {
                            var pLen = compPattern.Length;
                            var textLen = compText.Length;
                            int bestDist = int.MaxValue;
                            int bestPos = -1;
                            int maxWindow = Math.Min(2000, Math.Max(textLen, pLen + 20));
                            for (int i = 0; i < Math.Max(1, textLen - Math.Max(0, pLen - 1)); i++)
                            {
                                int len = Math.Min(pLen + fuzzyThreshold, textLen - i);
                                if (len <= 0) break;
                                string window = compText.Substring(i, len);
                                int dist = ClipDiscordApp.Utils.OcrTextUtils.LevenshteinDistance(window, compPattern);
                                if (dist < bestDist)
                                {
                                    bestDist = dist;
                                    bestPos = i;
                                    if (bestDist == 0) break;
                                }
                                if (i > maxWindow) break;
                            }

                            if (bestDist <= fuzzyThreshold)
                            {
                                var rawUpper = ClipDiscordApp.Utils.OcrTextUtils.NormalizeForComparison(rawText ?? string.Empty);
                                int rawIdx = Math.Max(0, Math.Min(rawUpper.Length - compPattern.Length, bestPos));
                                if (rawIdx >= 0 && rawIdx + compPattern.Length <= (rawText ?? "").Length)
                                {
                                    var matchedOriginal = (rawText ?? "").Substring(rawIdx, Math.Min(compPattern.Length, (rawText ?? "").Length - rawIdx));
                                    em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(matchedOriginal));
                                }
                                else
                                {
                                    em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(rawPattern));
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
                else // Regex タイプ
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
                                var rawUpper = ClipDiscordApp.Utils.OcrTextUtils.NormalizeForComparison(rawText ?? string.Empty);
                                int rawIdx = rawUpper.IndexOf(matchedComp, StringComparison.Ordinal);
                                if (rawIdx >= 0 && rawIdx + matchedComp.Length <= (rawText ?? "").Length)
                                {
                                    em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput((rawText ?? "").Substring(rawIdx, matchedComp.Length)));
                                }
                                else
                                {
                                    em.Matches.Add(ClipDiscordApp.Utils.OcrTextUtils.NormalizeForOutput(matchedComp));
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

            NextRule:
                continue;
            }

            System.Diagnostics.Debug.WriteLine($"[ParseByRules] Finished. totalMatches={results.Count}");
            return results;
        }
    }
}