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
        // デバッグ用フラグ: false にすると前処理をスキップして ROI をそのまま返す
        public static bool EnablePreprocessing = true;

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
        /// 前処理を無効化する場合は OcrPreprocessor.EnablePreprocessing = false にしてください。
        /// </summary>
        public static Bitmap Preprocess(Bitmap srcBmp, Params p = null)
        {
            p ??= new Params();
            if (srcBmp == null) throw new ArgumentNullException(nameof(srcBmp));

            // デバッグ／切り分け用: 前処理を無効化するフラグ
            if (!EnablePreprocessing)
            {
                return (Bitmap)srcBmp.Clone();
            }

            // Bitmap -> Mat
            using var srcMat = BitmapConverter.ToMat((Bitmap)srcBmp.Clone());

            Mat working = null;
            try
            {
                // リサイズ（Scale が 1 より大きい場合のみ）
                if (p.Scale > 1)
                {
                    working = new Mat();
                    Cv2.Resize(srcMat, working, new Size(srcMat.Width * p.Scale, srcMat.Height * p.Scale), 0, 0, InterpolationFlags.Cubic);
                }
                else
                {
                    working = srcMat.Clone();
                }

                // グレースケール化
                using var gray = new Mat();
                if (working.Channels() == 3)
                    Cv2.CvtColor(working, gray, ColorConversionCodes.BGR2GRAY);
                else if (working.Channels() == 4)
                    Cv2.CvtColor(working, gray, ColorConversionCodes.BGRA2GRAY);
                else
                    working.CopyTo(gray);

                // ノイズ除去（bilateral）を試行、失敗時は gray を使う
                using var denoised = new Mat();
                try
                {
                    Cv2.BilateralFilter(gray, denoised, Math.Max(1, p.BilateralDiameter), p.BilateralSigmaColor, p.BilateralSigmaSpace);
                }
                catch
                {
                    gray.CopyTo(denoised);
                }

                // EnhanceAndOutline を呼んで single-channel Mat を取得（caller が Dispose する）
                using var enhanced = EnhanceAndOutline(denoised, p);

                // ボーダー追加（クリッピング回避）
                var borderPx = Math.Max(0, p.Border);
                using var bordered = new Mat();
                Cv2.CopyMakeBorder(enhanced, bordered, borderPx, borderPx, borderPx, borderPx, BorderTypes.Constant, Scalar.All(255));

                // Mat -> Bitmap（Clone して Mat のライフサイクルから切り離す）
                var outBmp = BitmapConverter.ToBitmap(bordered.Clone());

                // デバッグ保存（オプション）
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
            Bitmap bmp = null;
            try
            {
                bmp = Preprocess(srcBmp, p);
                var pix = PixConverter.ToPix(bmp);
                bmp.Dispose();
                bmp = null;
                return pix;
            }
            catch
            {
                if (bmp != null) bmp.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Enhance and optionally outline. 入力は single-channel Mat（グレースケール）。
        /// 戻り値は caller が Dispose する必要がある Mat（single-channel, CV_8U）。
        /// </summary>
        private static Mat EnhanceAndOutline(Mat srcGray, Params p)
        {
            if (srcGray == null) throw new ArgumentNullException(nameof(srcGray));
            if (p == null) p = new Params();

            // work はメソッド内で所有するクローン
            var work = srcGray.Clone();

            Mat bw = null;
            Mat morph = null;
            Mat dilated = null;
            Mat finalMat = null;
            Mat resultMat = null;

            try
            {
                // 1) If dark background with white text, invert to white background / black text
                if (p.IsDarkBackground)
                {
                    Cv2.BitwiseNot(work, work);
                }

                // 2) Optional CLAHE / contrast
                if (p.UseClahe)
                {
                    var claheOut = new Mat();
                    try
                    {
                        using var clahe = Cv2.CreateCLAHE(p.ClaheClipLimit, p.ClaheTileGridSize);
                        clahe.Apply(work, claheOut);
                        work.Dispose();
                        work = claheOut;
                        claheOut = null;
                    }
                    catch
                    {
                        if (claheOut != null && !claheOut.IsDisposed) claheOut.Dispose();
                        // keep work as-is
                    }
                }

                // 3) small blur (ensure odd kernel >=1)
                var gk = p.GaussianKernel;
                if (gk <= 0) gk = 1;
                if (gk % 2 == 0) gk += 1;
                Cv2.GaussianBlur(work, work, new Size(gk, gk), 0);

                // 4) adaptive threshold (Binary)
                var blockSize = p.AdaptiveBlockSize;
                if (blockSize < 3) blockSize = 3;
                if (blockSize % 2 == 0) blockSize += 1;
                bw = new Mat();
                var threshType = p.InvertThreshold ? ThresholdTypes.BinaryInv : ThresholdTypes.Binary;
                try
                {
                    Cv2.AdaptiveThreshold(work, bw, 255, AdaptiveThresholdTypes.GaussianC, threshType, blockSize, p.AdaptiveC);
                }
                catch
                {
                    // fallback to Otsu
                    Cv2.Threshold(work, bw, 0, 255, ThresholdTypes.Binary | ThresholdTypes.Otsu);
                }

                // 5) Morph close -> open (fill holes then remove small noise)
                var kernelSize = Math.Max(1, p.MorphKernel);
                if (kernelSize % 2 == 0) kernelSize += 1;
                using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(kernelSize, kernelSize));
                morph = new Mat();
                try
                {
                    Cv2.MorphologyEx(bw, morph, MorphTypes.Close, kernel, iterations: 1);
                    Cv2.MorphologyEx(morph, morph, MorphTypes.Open, kernel, iterations: 1);
                }
                catch
                {
                    if (morph != null && !morph.IsDisposed) morph.Dispose();
                    morph = bw.Clone();
                }

                // 6) Dilate to thicken strokes (repair vertical gaps)
                Mat postMorph = morph;
                if (p.UseDilate)
                {
                    dilated = new Mat();
                    try
                    {
                        var dilateKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(2, 2));
                        Cv2.Dilate(morph, dilated, dilateKernel, iterations: Math.Max(1, p.DilateIterations));
                        postMorph = dilated;
                    }
                    catch
                    {
                        if (dilated != null && !dilated.IsDisposed) dilated.Dispose();
                        postMorph = morph;
                    }
                }

                // 7) Small close to fill micro gaps (3x3)
                using var smallKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3));
                try
                {
                    Cv2.MorphologyEx(postMorph, postMorph, MorphTypes.Close, smallKernel, iterations: 1);
                }
                catch
                {
                    // ignore and continue
                }

                // 8) Optional gentle sharpening (unsharp)
                finalMat = postMorph;
                Mat sharpened = null;
                if (p.Sharpen)
                {
                    try
                    {
                        using var blurForUnsharp = new Mat();
                        Cv2.GaussianBlur(postMorph, blurForUnsharp, new Size(3, 3), 0);
                        sharpened = new Mat();
                        Cv2.AddWeighted(postMorph, p.SharpenAmount, blurForUnsharp, p.SharpenBlurCoeff, 0, sharpened);
                        finalMat = sharpened;
                    }
                    catch
                    {
                        if (sharpened != null && !sharpened.IsDisposed) sharpened.Dispose();
                        finalMat = postMorph;
                    }
                }

                // 9) Optional outline (morphological gradient) and combine
                if (p.MakeOutline)
                {
                    using var gradKernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(3, 3));
                    using var outline = new Mat();
                    try
                    {
                        Cv2.MorphologyEx(finalMat, outline, MorphTypes.Gradient, gradKernel);

                        // Thicken outline to be visible
                        using var ok = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(2, 2));
                        Cv2.Dilate(outline, outline, ok, iterations: 1);

                        // Combine outline and filled text (bitwise OR) -> produce new Mat to return
                        resultMat = new Mat();
                        Cv2.BitwiseOr(finalMat, outline, resultMat);
                    }
                    catch
                    {
                        if (resultMat != null && !resultMat.IsDisposed) resultMat.Dispose();
                        resultMat = finalMat.Clone();
                    }
                }
                else
                {
                    resultMat = finalMat.Clone();
                }

                // Ensure single-channel CV_8U
                if (resultMat.Type() != MatType.CV_8U)
                {
                    var converted = new Mat();
                    resultMat.ConvertTo(converted, MatType.CV_8U);
                    resultMat.Dispose();
                    resultMat = converted;
                }

                return resultMat; // caller must Dispose
            }
            catch
            {
                if (bw != null && !bw.IsDisposed) bw.Dispose();
                if (morph != null && !morph.IsDisposed) morph.Dispose();
                if (dilated != null && !dilated.IsDisposed) dilated.Dispose();
                if (finalMat != null && !finalMat.IsDisposed && finalMat != morph && finalMat != dilated) finalMat.Dispose();
                if (resultMat != null && !resultMat.IsDisposed) resultMat.Dispose();
                if (work != null && !work.IsDisposed) work.Dispose();
                throw;
            }
            finally
            {
                if (bw != null && !bw.IsDisposed) bw.Dispose();
                if (morph != null && !morph.IsDisposed) morph.Dispose();
                if (dilated != null && !dilated.IsDisposed) dilated.Dispose();
                if (finalMat != null && !finalMat.IsDisposed && finalMat != resultMat) finalMat.Dispose();
                if (work != null && !work.IsDisposed) work.Dispose();
            }
        }

        // 既存クラスがある前提。なければ別ファイルで同名 static クラスを作る。
        public static Bitmap GetGrayBitmap(Bitmap prepped)
        {
            if (prepped == null) return null;
            var bmp = new Bitmap(prepped.Width, prepped.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            using (var g = Graphics.FromImage(bmp))
            {
                var cm = new System.Drawing.Imaging.ColorMatrix(new float[][]
                {
                    new float[] {0.299f, 0.299f, 0.299f, 0, 0},
                    new float[] {0.587f, 0.587f, 0.587f, 0, 0},
                    new float[] {0.114f, 0.114f, 0.114f, 0, 0},
                    new float[] {0,      0,      0,      1, 0},
                    new float[] {0,      0,      0,      0, 1}
                });
                var ia = new System.Drawing.Imaging.ImageAttributes();
                ia.SetColorMatrix(cm);
                g.DrawImage(prepped, new Rectangle(0, 0, prepped.Width, prepped.Height),
                    0, 0, prepped.Width, prepped.Height, GraphicsUnit.Pixel, ia);
            }
            return bmp;
        }
    }
}