using System;

namespace ClipDiscordApp.Models
{
    public enum ExtractRuleType
    {
        Regex,
        Keyword
    }

    public class ExtractRule
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = "";
        public string Pattern { get; set; } = "";
        public ExtractRuleType Type { get; set; } = ExtractRuleType.Regex;
        public bool Enabled { get; set; } = true;
        public int Order { get; set; } = 0;
    }
}