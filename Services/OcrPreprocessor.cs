using System;
using System.Drawing;
using System.Drawing.Imaging;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Tesseract;
using Size = OpenCvSharp.Size;

namespace YourNamespace.Ocr
{
    public static class OcrPreprocessor
    {
        public class Params
        {
            public int Scale = 2;
            public int BilateralDiameter = 9;
            public int BilateralSigmaColor = 75;
            public int BilateralSigmaSpace = 75;
            public int GaussianKernel = 3;
            public int AdaptiveBlockSize = 31;
            public double AdaptiveC = 10.0;
            public int MorphKernel = 2;
            public bool UseClahe = true;
            public double ClaheClipLimit = 3.0;
            public Size ClaheTileGridSize = new Size(8, 8);
            public bool Sharpen = false;
        }

        public static Bitmap Preprocess(Bitmap srcBmp, Params p = null)
        {
            p ??= new Params();
            if (srcBmp == null) throw new ArgumentNullException(nameof(srcBmp));

            using var srcMat = BitmapConverter.ToMat(srcBmp);

            // 拡大 or コピー（new Mat(Size, MatType) を使う）
            Mat mat;
            if (p.Scale > 1)
            {
                mat = new Mat();
                Cv2.Resize(srcMat, mat, new Size(srcMat.Width * p.Scale, srcMat.Height * p.Scale), 0, 0, InterpolationFlags.Cubic);
            }
            else
            {
                mat = srcMat.Clone();
            }

            try
            {
                // グレースケール
                using var gray = new Mat();
                if (mat.Channels() == 3)
                    Cv2.CvtColor(mat, gray, ColorConversionCodes.BGR2GRAY);
                else if (mat.Channels() == 4)
                    Cv2.CvtColor(mat, gray, ColorConversionCodes.BGRA2GRAY);
                else
                    mat.CopyTo(gray);

                // ノイズ低減（bilateral）
                using var denoised = new Mat();
                Cv2.BilateralFilter(gray, denoised, p.BilateralDiameter, p.BilateralSigmaColor, p.BilateralSigmaSpace);

                // CLAHE
                Mat contrastMat = denoised;
                Mat tempContrast = null;
                if (p.UseClahe)
                {
                    tempContrast = new Mat();
                    using var clahe = Cv2.CreateCLAHE(p.ClaheClipLimit, p.ClaheTileGridSize);
                    clahe.Apply(denoised, tempContrast);
                    contrastMat = tempContrast;
                }

                // ガウスぼかし
                using var smooth = new Mat();
                Cv2.GaussianBlur(contrastMat, smooth, new Size(p.GaussianKernel, p.GaussianKernel), 0);

                // 適応二値化（位置引数）
                using var bw = new Mat();
                Cv2.AdaptiveThreshold(smooth, bw, 255, AdaptiveThresholdTypes.GaussianC, ThresholdTypes.Binary, p.AdaptiveBlockSize, (int)p.AdaptiveC);

                // 形態学的処理
                using var kernel = Cv2.GetStructuringElement(MorphShapes.Rect, new Size(p.MorphKernel, p.MorphKernel));
                using var morph = new Mat();
                Cv2.MorphologyEx(bw, morph, MorphTypes.Open, kernel);
                Cv2.MorphologyEx(morph, morph, MorphTypes.Close, kernel);

                // シャープ化（任意）
                Mat finalMat = morph;
                Mat sharpened = null;
                if (p.Sharpen)
                {
                    // new Mat(new Size, MatType) を使用して安全に作成
                    var k = new Mat(new Size(3, 3), MatType.CV_32F);
                    // カーネル値を Set で入れる（Mat は確実に Mat 型）
                    k.Set<float>(0, 0, 0);
                    k.Set<float>(0, 1, -1);
                    k.Set<float>(0, 2, 0);
                    k.Set<float>(1, 0, -1);
                    k.Set<float>(1, 1, 5);
                    k.Set<float>(1, 2, -1);
                    k.Set<float>(2, 0, 0);
                    k.Set<float>(2, 1, -1);
                    k.Set<float>(2, 2, 0);

                    sharpened = new Mat();
                    Cv2.Filter2D(morph, sharpened, MatType.CV_8U, k);
                    k.Dispose();

                    finalMat = sharpened;
                }

                // 出力に1ピクセル余白を追加
                using var bordered = new Mat();
                Cv2.CopyMakeBorder(finalMat, bordered, 1, 1, 1, 1, BorderTypes.Constant, Scalar.All(255));

                // Bitmap に戻す（Clone して返す）
                var outBmp = BitmapConverter.ToBitmap(bordered.Clone());

                if (tempContrast != null) tempContrast.Dispose();
                if (sharpened != null) sharpened.Dispose();
                mat.Dispose();

                return outBmp;
            }
            finally
            {
                if (!mat.IsDisposed) mat.Dispose();
            }
        }

        // Pix 生成は PixConverter に任せる（存在しない場合は明確な例外を投げる）
        public static Pix PreprocessToPix(Bitmap srcBmp, Params p = null)
        {
            var bmp = Preprocess(srcBmp, p);
            try
            {
                // PixConverter が利用可能な場合はこれで変換
                var pix = PixConverter.ToPix(bmp);
                bmp.Dispose();
                return pix;
            }
            catch (MissingMethodException)
            {
                bmp.Dispose();
                throw new InvalidOperationException("PixConverter.ToPix が利用できません。Tesseract パッケージのバージョンを確認するか、呼び出し側で Bitmap->Pix 変換を実装してください。");
            }
            catch (Exception ex)
            {
                bmp.Dispose();
                throw new InvalidOperationException("Pix 変換に失敗しました: " + ex.Message, ex);
            }
        }
    }
}