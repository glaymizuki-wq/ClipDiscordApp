using System;
using System.Drawing;
using System.IO;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Tesseract;
using Size = OpenCvSharp.Size;

namespace YourNamespace.Ocr
{
    /// <summary>
    /// 画像前処理ユーティリティ（Tesseract 用 Pix 変換も提供）
    /// </summary>
    public static class OcrPreprocessor
    {
        public class Params
        {
            // Resize / denoise / blur
            public int Scale = 4;
            public int BilateralDiameter = 9;
            public int BilateralSigmaColor = 75;
            public int BilateralSigmaSpace = 75;
            public int GaussianKernel = 3;

            // Adaptive threshold
            public int AdaptiveBlockSize = 31; // must be odd
            public double AdaptiveC = 8.0;
            public bool InvertThreshold = false;

            // CLAHE
            public bool UseClahe = true;
            public double ClaheClipLimit = 3.0;
            public Size ClaheTileGridSize = new Size(8, 8);

            // Morphology / dilation / border
            public int MorphKernel = 7;
            public bool UseDilate = true;
            public int DilateIterations = 1;
            public int Border = 8;

            // Sharpen control
            public bool Sharpen = false;
            public double SharpenAmount = 1.2;
            public double SharpenBlurCoeff = -0.2;

            // Background / outline options
            public bool IsDarkBackground = false; // true = black background / white text
            public bool MakeOutline = false;      // generate outline and OR with text

            // Debug: save preprocessed PNG
            public bool SavePreprocessed = false;
            public string PreprocessedSaveFolder = null; // if null, uses executing dir
            public string PreprocessedFilenamePrefix = "pre";
        }

        /// <summary>
        /// 前処理を実行し、Bitmap を返します。caller は返り値 Bitmap を Dispose してください。
        /// </summary>
        public static Bitmap Preprocess(Bitmap srcBmp, Params p = null)
        {
            p ??= new Params();
            if (srcBmp == null) throw new ArgumentNullException(nameof(srcBmp));

            // Convert Bitmap -> Mat
            using var srcMat = BitmapConverter.ToMat((Bitmap)srcBmp.Clone());

            // Resize if required
            Mat working = null;
            if (p.Scale > 1)
            {
                working = new Mat();
                Cv2.Resize(srcMat, working, new Size(srcMat.Width * p.Scale, srcMat.Height * p.Scale), 0, 0, InterpolationFlags.Cubic);
            }
            else
            {
                working = srcMat.Clone();
            }

            try
            {
                // 1) To Gray
                using var gray = new Mat();
                if (working.Channels() == 3)
                    Cv2.CvtColor(working, gray, ColorConversionCodes.BGR2GRAY);
                else if (working.Channels() == 4)
                    Cv2.CvtColor(working, gray, ColorConversionCodes.BGRA2GRAY);
                else
                    working.CopyTo(gray);

                // 2) Denoise (bilateral) - keeps edges
                using var denoised = new Mat();
                Cv2.BilateralFilter(gray, denoised, p.BilateralDiameter, p.BilateralSigmaColor, p.BilateralSigmaSpace);

                // 3) Enhance, outline and obtain final single-channel Mat
                // NOTE: EnhanceAndOutline returns a NEW Mat which caller must Dispose
                using var enhanced = EnhanceAndOutline(denoised, p);

                // 4) Add border to avoid clipping (use configured border)
                var borderPx = Math.Max(0, p.Border);
                using var bordered = new Mat();
                Cv2.CopyMakeBorder(enhanced, bordered, borderPx, borderPx, borderPx, borderPx, BorderTypes.Constant, Scalar.All(255));

                // 5) Convert to Bitmap
                var outBmp = BitmapConverter.ToBitmap(bordered.Clone());

                // 6) Optionally save preprocessed PNG for debugging
                if (p.SavePreprocessed)
                {
                    try
                    {
                        var folder = p.PreprocessedSaveFolder;
                        if (string.IsNullOrEmpty(folder)) folder = AppDomain.CurrentDomain.BaseDirectory;
                        if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                        var timestamp = DateTime.UtcNow.ToString("yyyyMMdd_HHmmss_fff");
                        var filename = $"{p.PreprocessedFilenamePrefix}_{timestamp}.png";
                        var full = Path.Combine(folder, filename);
                        outBmp.Save(full, System.Drawing.Imaging.ImageFormat.Png);
                        System.Diagnostics.Debug.WriteLine($"[OcrPreprocessor] Preprocessed image saved: {full}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OcrPreprocessor] Failed to save preprocessed image: {ex}");
                    }
                }

                return outBmp;
            }
            finally
            {
                if (working != null && !working.IsDisposed) working.Dispose();
            }
        }

        /// <summary>
        /// Preprocess して Pix に変換して返す。呼び出し側は返却 Pix を Dispose すること。
        /// </summary>
        public static Pix PreprocessToPix(Bitmap srcBmp, Params p = null)
        {
            var bmp = Preprocess(srcBmp, p);
            try
            {
                var pix = PixConverter.ToPix(bmp);
                bmp.Dispose();
                return pix;
            }
            catch
            {
                bmp.Dispose();
                throw;
            }
        }

        // EnhanceAndOutline メソッド（差し替え用 完全版）
        // 入力: 単一チャネルの Mat (グレースケール)
        // 出力: 新たに生成された Mat（caller が Dispose すること）
        private static Mat EnhanceAndOutline(Mat srcGray, Params p)
        {
            if (srcGray == null) throw new ArgumentNullException(nameof(srcGray));

            // work はメソッド内で所有するクローン
            var work = srcGray.Clone();

            try
            {
                // 1) If dark background with white text, invert to white background / black text
                if (p.IsDarkBackground)
                {
                    Cv2.BitwiseNot(work, work);
                }

                // 2) Optional CLAHE / contrast
                Mat temp = null;
                if (p.UseClahe)
                {
                    temp = new Mat();
                    using (var clahe = Cv2.CreateCLAHE(p.ClaheClipLimit, p.ClaheTileGridSize))
                    {
                        clahe.Apply(work, temp);
                    }
                    work.Dispose();
                    work = temp;
                    temp = null;
                }

                // 3) small blur
                var gk = p.GaussianKernel;
                if (gk <= 0) gk = 1;
                if (gk % 2 == 0) gk += 1;
                Cv2.GaussianBlur(work, work, new Size(gk, gk), 0);

                // 4) adaptive threshold (Binary)
                var blockSize = p.AdaptiveBlockSize;
                if (blockSize < 3) blockSize = 3;
                if (blockSize % 2 == 0) blockSize += 1;
                var bw = new Mat();
                var threshType = p.InvertThreshold ? ThresholdTypes.BinaryInv : ThresholdTypes.Binary;
                Cv2.AdaptiveThreshold(work, bw, 255, AdaptiveThresholdTypes.GaussianC, threshType, blockSize, p.AdaptiveC);

                // 5) Morph close -> open (fill holes then remove small noise)
                var kernelSize = Math.Max(1, p.MorphKernel);
                using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(kernelSize, kernelSize));
                var morph = new Mat();
                Cv2.MorphologyEx(bw, morph, MorphTypes.Close, kernel, iterations: 1);
                Cv2.MorphologyEx(morph, morph, MorphTypes.Open, kernel, iterations: 1);

                // 6) Dilate to thicken strokes (repair vertical gaps)
                Mat postMorph = morph;
                Mat dilated = null;
                if (p.UseDilate)
                {
                    dilated = new Mat();
                    var dilateKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(2, 2));
                    Cv2.Dilate(morph, dilated, dilateKernel, iterations: Math.Max(1, p.DilateIterations));
                    postMorph = dilated;
                }

                // 7) Small close to fill micro gaps (3x3)
                using var smallKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3));
                Cv2.MorphologyEx(postMorph, postMorph, MorphTypes.Close, smallKernel, iterations: 1);

                // 8) Optional gentle sharpening (unsharp)
                Mat finalMat = postMorph;
                Mat sharpened = null;
                if (p.Sharpen)
                {
                    using var blurForUnsharp = new Mat();
                    Cv2.GaussianBlur(postMorph, blurForUnsharp, new Size(3, 3), 0);
                    sharpened = new Mat();
                    Cv2.AddWeighted(postMorph, p.SharpenAmount, blurForUnsharp, p.SharpenBlurCoeff, 0, sharpened);
                    finalMat = sharpened;
                }

                // 9) Optional outline (morphological gradient) and combine
                Mat resultMat;
                if (p.MakeOutline)
                {
                    using var gradKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3));
                    using var outline = new Mat();
                    Cv2.MorphologyEx(finalMat, outline, MorphTypes.Gradient, gradKernel);

                    // Thicken outline to be visible
                    using var ok = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(2, 2));
                    Cv2.Dilate(outline, outline, ok, iterations: 1);

                    // Combine outline and filled text (bitwise OR) -> produce new Mat to return
                    resultMat = new Mat();
                    Cv2.BitwiseOr(finalMat, outline, resultMat);
                }
                else
                {
                    // Return a clone to ensure caller owns returned Mat
                    resultMat = finalMat.Clone();
                }

                // cleanup temporaries
                if (bw != null && !bw.IsDisposed) bw.Dispose();
                if (morph != null && !morph.IsDisposed) morph.Dispose();
                if (dilated != null && !dilated.IsDisposed) dilated.Dispose();
                if (sharpened != null && !sharpened.IsDisposed) sharpened.Dispose();
                if (work != null && !work.IsDisposed) work.Dispose();

                return resultMat; // caller must Dispose
            }
            catch
            {
                if (work != null && !work.IsDisposed) work.Dispose();
                throw;
            }
        }
    }
}