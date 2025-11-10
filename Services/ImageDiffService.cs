using System;
using System.Drawing;

namespace ClipDiscordApp.Services
{
    public class ImageDiffService
    {
        public bool HasSignificantChange(Bitmap prev, Bitmap current, double threshold = 0.03)
        {
            if (prev == null || current == null) return true;
            if (prev.Width != current.Width || prev.Height != current.Height) return true;

            int width = prev.Width;
            int height = prev.Height;
            long totalDiff = 0;
            long maxDiff = width * height * 3 * 255;

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    var c1 = prev.GetPixel(x, y);
                    var c2 = current.GetPixel(x, y);

                    totalDiff += Math.Abs(c1.R - c2.R);
                    totalDiff += Math.Abs(c1.G - c2.G);
                    totalDiff += Math.Abs(c1.B - c2.B);
                }
            }

            double diffRatio = (double)totalDiff / maxDiff;
            return diffRatio >= threshold;
        }
    }
}