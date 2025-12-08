namespace ClipDiscordApp.Models
{
    // MainWindow 等で参照している簡易ラッパー型
    public class MatchResultItem
    {
        public string RuleId { get; set; }
        public string RuleName { get; set; }
        public string[] Matches { get; set; }

        public MatchResultItem() { }

        public MatchResultItem(string ruleId, string ruleName, string[] matches)
        {
            RuleId = ruleId;
            RuleName = ruleName;
            Matches = matches ?? new string[0];
        }
    }
}