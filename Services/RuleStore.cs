using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ClipDiscordApp.Models;

namespace ClipDiscordApp.Services
{
    public static class RuleStore
    {
        private static readonly string PathRules = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rules.json");

        public static List<ExtractRule> Load()
        {
            try
            {
                if (!File.Exists(PathRules)) return new List<ExtractRule>();
                var json = File.ReadAllText(PathRules);
                var opts = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var list = JsonSerializer.Deserialize<List<ExtractRule>>(json, opts) ?? new List<ExtractRule>();
                list.Sort((a, b) => a.Order.CompareTo(b.Order));
                return list;
            }
            catch
            {
                return new List<ExtractRule>();
            }
        }

        public static void Save(IEnumerable<ExtractRule> rules)
        {
            var list = new List<ExtractRule>(rules);
            list.Sort((a, b) => a.Order.CompareTo(b.Order));
            var opts = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(list, opts);
            File.WriteAllText(PathRules, json);
        }
    }
}