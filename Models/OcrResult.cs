using System;
using System.Collections.Generic;
using System.Drawing;

namespace ClipDiscordApp.Models
{
    public class OcrWord
    {
        public string Text { get; set; } = "";
        public float Confidence { get; set; }
        public System.Drawing.Rectangle BoundingBox { get; set; }
    }

    public class OcrResult
    {
        public DateTime TimestampUtc { get; set; }
        public Rectangle Region { get; set; }
        public string FullText { get; set; } = "";
        public List<OcrWord> Words { get; set; } = new();
        public string RawText => string.Join(" ", Words.Select(w => w.Text));

    }
}