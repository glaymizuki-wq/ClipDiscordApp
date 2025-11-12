using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace ClipDiscordApp.Utils
{
    /// <summary>
    /// OCR 用キャプチャ／前処理補助
    /// - 返却する Bitmap の Dispose は呼び出し側の責任です
    /// - 保存は例外を吸収するベストエフォート
    /// </summary>
    public static class OcrCaptureHelper
    {
        /// <summary>
        /// 画面の指定領域をキャプチャして Bitmap を返す。
        /// scale &gt; 1 のときは高品質に拡大して返す。
        /// </summary>
        public static Bitmap CaptureScreenRegion(Rectangle region, int scale = 1)
        {
            if (region.Width <= 0 || region.Height <= 0) throw new ArgumentException("region must be positive");

            var capture = new Bitmap(region.Width, region.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(capture))
            {
                g.CopyFromScreen(region.Left, region.Top, 0, 0, region.Size, CopyPixelOperation.SourceCopy);
            }

            if (scale <= 1) return capture;

            var scaled = new Bitmap(region.Width * scale, region.Height * scale, PixelFormat.Format32bppArgb);
            using (var g2 = Graphics.FromImage(scaled))
            {
                g2.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g2.DrawImage(capture, 0, 0, scaled.Width, scaled.Height);
            }
            capture.Dispose();
            return scaled;
        }

        /// <summary>
        /// 右端ラベル行を切り出す（返却 Bitmap は呼び出し側で Dispose）。
        /// rightPct: 切り出す幅の割合（例 0.32）。yCenterPct/heightPct は相対位置。
        /// </summary>
        public static Bitmap CropLabelRow(Bitmap src, double rightPct = 0.32, double yCenterPct = 0.12, double heightPct = 0.12)
        {
            if (src == null) throw new ArgumentNullException(nameof(src));
            rightPct = Clamp01(rightPct);
            yCenterPct = Clamp01(yCenterPct);
            heightPct = Clamp01(heightPct);

            int w = src.Width, h = src.Height;
            int cropW = Math.Max(1, (int)(w * rightPct));
            int centerY = (int)(h * yCenterPct);
            int cropH = Math.Max(1, (int)(h * heightPct));
            int y = Math.Max(0, Math.Min(h - 1, centerY - cropH / 2));
            if (y + cropH > h) cropH = h - y;

            var rect = new Rectangle(w - cropW, y, cropW, cropH);
            return src.Clone(rect, src.PixelFormat);
        }

        /// <summary>
        /// 生画像を保存して（オプション）平均輝度で反転が必要か判定し、
        /// 必要なら反転した Bitmap を返す。元 bmp は内部で Dispose される（呼び出し側に返す Bitmap を破棄する責任あり）。
        /// </summary>
        public static Bitmap PrecheckAndFix(Bitmap bmp, string saveFolder = null, int invertThreshold = 110)
        {
            if (bmp == null) return null;

            // Save raw for debugging (best-effort)
            try
            {
                if (!string.IsNullOrEmpty(saveFolder))
                {
                    Directory.CreateDirectory(saveFolder);
                    var p = Path.Combine(saveFolder, $"raw_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                    bmp.Save(p, ImageFormat.Png);
                }
            }
            catch { /* ignore save errors */ }

            // compute average brightness using LockBits -> byte array for speed
            double avg = ComputeAverageBrightness(bmp);

            // If image is dark (black background / white text), invert to white background / black text
            if (avg < invertThreshold)
            {
                var inv = new Bitmap(bmp.Width, bmp.Height, bmp.PixelFormat);
                using (var g = Graphics.FromImage(inv))
                {
                    var cm = new ColorMatrix(new float[][]
                    {
                        new float[]{-1,0,0,0,0},
                        new float[]{0,-1,0,0,0},
                        new float[]{0,0,-1,0,0},
                        new float[]{0,0,0,1,0},
                        new float[]{1,1,1,0,1}
                    });
                    using var ia = new ImageAttributes();
                    ia.SetColorMatrix(cm);
                    g.DrawImage(bmp, new Rectangle(0, 0, bmp.Width, bmp.Height),
                        0, 0, bmp.Width, bmp.Height, GraphicsUnit.Pixel, ia);
                }
                bmp.Dispose();
                return inv;
            }

            // no change
            return bmp;
        }

        /// <summary>
        /// 指定パラメータ群からリトライ候補画像を遅延列挙で返す。
        /// 列挙するときに Bitmap を生成するため、foreach の各要素を使い終わったら Dispose してください。
        /// </summary>
        public static IEnumerable<Bitmap> GenerateRetryCandidates(Bitmap original, IEnumerable<int> scales, IEnumerable<double> rightPcts, IEnumerable<double> yShifts)
        {
            if (original == null) yield break;
            var scaleList = (scales ?? new[] { 1 }).Distinct().ToArray();
            var rpList = (rightPcts ?? new[] { 0.32 }).Distinct().ToArray();
            var ysList = (yShifts ?? new[] { 0.0 }).Distinct().ToArray();

            foreach (var s in scaleList)
            {
                Bitmap scaled = null;
                try
                {
                    if (s <= 1)
                        scaled = (Bitmap)original.Clone();
                    else
                        scaled = ResizeBitmap((Bitmap)original.Clone(), s);

                    foreach (var rp in rpList)
                    {
                        foreach (var ys in ysList)
                        {
                            Bitmap candidate = null;
                            try
                            {
                                var yCenterPct = 0.12 + ys;
                                candidate = CropLabelRow(scaled, rp, yCenterPct, 0.12);
                                yield return candidate;
                            }
                            finally
                            {
                                // the yielded candidate must be disposed by the caller; do not dispose here
                                // if candidate was not yielded due to exception, ensure it's disposed
                                // (but in normal flow candidate is yielded)
                                candidate = null;
                            }
                        }
                    }
                }
                finally
                {
                    // scaled should be disposed here if we created it (but not when we cloned and intended to crop clones)
                    scaled?.Dispose();
                }
            }
        }

        /// <summary>
        /// デバッグ用に一意ファイル名で保存する（best-effort）。
        /// </summary>
        public static void SaveDebugImage(Bitmap bmp, string folder, string prefix)
        {
            if (bmp == null) return;
            try
            {
                Directory.CreateDirectory(folder);
                var fname = Path.Combine(folder, $"{prefix}_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                bmp.Save(fname, ImageFormat.Png);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OcrCaptureHelper] SaveDebugImage failed: {ex.Message}");
            }
        }

        /// <summary>
        /// 黒背景判定（平均輝度）。閾値は 0..255。
        /// </summary>
        public static bool IsDarkBackground(Bitmap bmp, int threshold = 110)
        {
            if (bmp == null) return false;
            return ComputeAverageBrightness(bmp) < threshold;
        }

        // ----------------- internal helpers -----------------

        private static double ComputeAverageBrightness(Bitmap bmp)
        {
            var rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            var pf = bmp.PixelFormat;
            BitmapData data = null;
            try
            {
                data = bmp.LockBits(rect, ImageLockMode.ReadOnly, pf);
                int bytes = Math.Abs(data.Stride) * bmp.Height;
                var buffer = new byte[bytes];
                System.Runtime.InteropServices.Marshal.Copy(data.Scan0, buffer, 0, bytes);

                int bpp = Image.GetPixelFormatSize(pf) / 8;
                if (bpp < 3) // grayscale or indexed
                {
                    long sum = 0;
                    for (int i = 0; i < buffer.Length; i++)
                    {
                        sum += buffer[i];
                    }
                    bmp.UnlockBits(data);
                    return (double)sum / buffer.Length;
                }

                long total = 0;
                long count = 0;
                for (int y = 0; y < bmp.Height; y++)
                {
                    var rowStart = y * data.Stride;
                    for (int x = 0; x < bmp.Width; x++)
                    {
                        int idx = rowStart + x * bpp;
                        if (idx + 2 >= buffer.Length) break;
                        byte b = buffer[idx + 0];
                        byte g = buffer[idx + 1];
                        byte r = buffer[idx + 2];
                        total += (r + g + b) / 3;
                        count++;
                    }
                }
                bmp.UnlockBits(data);
                return count == 0 ? 0.0 : (double)total / count;
            }
            finally
            {
                try { if (data != null) bmp.UnlockBits(data); } catch { }
            }
        }

        private static Bitmap ResizeBitmap(Bitmap src, int scale)
        {
            var scaled = new Bitmap(src.Width * scale, src.Height * scale, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(scaled))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(src, 0, 0, scaled.Width, scaled.Height);
            }
            src.Dispose();
            return scaled;
        }

        private static double Clamp01(double v) => v < 0 ? 0 : (v > 1 ? 1 : v);
    }
}