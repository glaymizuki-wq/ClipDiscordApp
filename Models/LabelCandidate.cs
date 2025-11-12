using System;

namespace ClipDiscordApp.Models
{
    public class LabelCandidate
    {
        public string Text { get; set; }
        public string Label { get; set; } // "BUY" or "SELL"
        public double Confidence { get; set; } // 0.0 .. 1.0
        public string Source { get; set; } // "binary"/"gray"/"template"/"psm=..."

        public LabelCandidate() { }

        public LabelCandidate(string text, string label, double confidence, string source = null)
        {
            Text = text;
            Label = label;
            Confidence = confidence;
            Source = source;
        }

        public override string ToString()
        {
            return $"{Label}:{Text} (conf={Confidence:F2}, src={Source})";
        }
    }
}