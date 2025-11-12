using System.Drawing;
using System.Drawing.Imaging;

namespace ClipDiscordApp.Services
{
    public static class OcrPreprocessorExtensions
    {
        public static Bitmap GetGrayBitmap(Bitmap prepped)
        {
            if (prepped == null) return null;
            var bmp = new Bitmap(prepped.Width, prepped.Height, PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bmp))
            {
                var cm = new ColorMatrix(new float[][]
                {
                    new float[] {0.299f, 0.299f, 0.299f, 0, 0},
                    new float[] {0.587f, 0.587f, 0.587f, 0, 0},
                    new float[] {0.114f, 0.114f, 0.114f, 0, 0},
                    new float[] {0,      0,      0,      1, 0},
                    new float[] {0,      0,      0,      0, 1}
                });
                var ia = new ImageAttributes();
                ia.SetColorMatrix(cm);
                g.DrawImage(prepped, new Rectangle(0, 0, prepped.Width, prepped.Height),
                    0, 0, prepped.Width, prepped.Height, GraphicsUnit.Pixel, ia);
            }
            return bmp;
        }

        // 簡易二値化（デバッグ用）。必要なら OpenCV 版に置換。
        public static Bitmap GetBinaryBitmap(Bitmap prepped)
        {
            if (prepped == null) return null;
            using var gray = GetGrayBitmap(prepped);
            var bw = new Bitmap(gray.Width, gray.Height, PixelFormat.Format24bppRgb);
            for (int y = 0; y < gray.Height; y++)
            {
                for (int x = 0; x < gray.Width; x++)
                {
                    var c = gray.GetPixel(x, y);
                    var lum = (int)(0.299 * c.R + 0.587 * c.G + 0.114 * c.B);
                    var v = lum > 128 ? 255 : 0;
                    bw.SetPixel(x, y, Color.FromArgb(v, v, v));
                }
            }
            return bw;
        }
    }
}