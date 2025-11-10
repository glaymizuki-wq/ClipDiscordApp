using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClipDiscordApp.Utils
{    public static class ImageHelper
    {
        public static bool IsImageDifferent(Bitmap img1, Bitmap img2, double threshold = 0.01)
        {
            if (img1.Width != img2.Width || img1.Height != img2.Height)
                return true;

            int diffCount = 0;
            int totalPixels = img1.Width * img1.Height;

            for (int y = 0; y < img1.Height; y++)
            {
                for (int x = 0; x < img1.Width; x++)
                {
                    Color c1 = img1.GetPixel(x, y);
                    Color c2 = img2.GetPixel(x, y);

                    if (Math.Abs(c1.R - c2.R) > 10 ||
                        Math.Abs(c1.G - c2.G) > 10 ||
                        Math.Abs(c1.B - c2.B) > 10)
                    {
                        diffCount++;
                    }
                }
            }

            double diffRatio = (double)diffCount / totalPixels;
            return diffRatio > threshold;
        }

        public static bool IsImageDifferentFast(Bitmap bmp1, Bitmap bmp2, double threshold = 0.01)
        {
            if (bmp1.Width != bmp2.Width || bmp1.Height != bmp2.Height)
                return true;

            int width = bmp1.Width;
            int height = bmp1.Height;
            int totalPixels = width * height;
            int diffCount = 0;

            var rect = new Rectangle(0, 0, width, height);
            var data1 = bmp1.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            var data2 = bmp2.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

            try
            {
                unsafe
                {
                    byte* ptr1 = (byte*)data1.Scan0;
                    byte* ptr2 = (byte*)data2.Scan0;
                    int stride = data1.Stride;

                    for (int y = 0; y < height; y++)
                    {
                        byte* row1 = ptr1 + (y * stride);
                        byte* row2 = ptr2 + (y * stride);

                        for (int x = 0; x < width; x++)
                        {
                            int i = x * 3;
                            int bDiff = Math.Abs(row1[i] - row2[i]);
                            int gDiff = Math.Abs(row1[i + 1] - row2[i + 1]);
                            int rDiff = Math.Abs(row1[i + 2] - row2[i + 2]);

                            if (rDiff > 10 || gDiff > 10 || bDiff > 10)
                                diffCount++;
                        }
                    }
                }
            }
            finally
            {
                bmp1.UnlockBits(data1);
                bmp2.UnlockBits(data2);
            }

            double diffRatio = (double)diffCount / totalPixels;
            return diffRatio > threshold;
        }
    }
}
