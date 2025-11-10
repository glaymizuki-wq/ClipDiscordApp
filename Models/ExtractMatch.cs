using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClipDiscordApp.Models
{
    public class ExtractMatch
    {
        // ルール識別子（ルールを追跡するための ID）
        public string RuleId { get; set; }

        // ルール名（ログや UI 表示用）
        public string RuleName { get; set; }

        // マッチした箇所（原文の切り出し・正規化済み文字列を複数持てる）
        public List<string> Matches { get; set; } = new List<string>();

        // 追加情報（必要なメタデータを格納できる柔軟フィールド）
        public Dictionary<string, string> Metadata { get; set; } = new Dictionary<string, string>();

        // マッチの信頼度（0.0〜1.0）や内部で使う評価値を入れると便利
        public double Confidence { get; set; } = 0.0;
    }
}
