using ClipDiscordApp.Models;
using ClipDiscordApp.Parsers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tesseract;
using YourNamespace.Ocr;
using DrawingImageFormat = System.Drawing.Imaging.ImageFormat;

namespace ClipDiscordApp.Services
{
    public class OcrService : IDisposable
    {
        private readonly TesseractEngine _engine;
        private bool _disposed;

        public OcrService(string tessdataPath, string lang = "jpn")
        {
            _engine = new TesseractEngine(tessdataPath, lang, EngineMode.Default);

            // Whitelist と PSM の設定（SingleWord を推奨）
            try
            {
                _engine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789:- ");
                // PSM 設定はラッパーのバージョンによってセット方法が異なる場合があるが、
                // 多くのラッパーでは以下で設定可能
                _engine.DefaultPageSegMode = PageSegMode.SingleWord;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OcrService] Failed to set tess variables: {ex}");
            }
        }

        /// <summary>
        /// 既存の低レベル OCR。Bitmap -> OcrResult（words, fullText）を返す。
        /// 変更なしで従来の用途に使用可能。
        /// </summary>
        public async Task<OcrResult> RecognizeAsync(Bitmap bmp, Rectangle captureRegion, CancellationToken ct = default)
        {
            if (bmp == null) throw new ArgumentNullException(nameof(bmp));

            return await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();

                var result = new OcrResult
                {
                    TimestampUtc = DateTime.UtcNow,
                    Region = captureRegion
                };

                using var pix = ConvertBitmapToPix(bmp);
                using var page = _engine.Process(pix);
                result.FullText = page.GetText()?.Trim() ?? "";

                using var iter = page.GetIterator();
                iter.Begin();

                do
                {
                    string text = iter.GetText(PageIteratorLevel.Word);
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    float confidence = iter.GetConfidence(PageIteratorLevel.Word);
                    if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out var box))
                    {
                        var r = new Rectangle(box.X1, box.Y1, box.Width, box.Height);
                        result.Words.Add(new OcrWord
                        {
                            Text = text.Trim(),
                            Confidence = confidence,
                            BoundingBox = r
                        });
                    }
                    else
                    {
                        result.Words.Add(new OcrWord
                        {
                            Text = text.Trim(),
                            Confidence = confidence,
                            BoundingBox = new Rectangle(0, 0, 0, 0)
                        });
                    }
                } while (iter.Next(PageIteratorLevel.Word));

                return result;
            }, ct);
        }

        private Pix ConvertBitmapToPix(Bitmap bmp)
        {
            using var ms = new MemoryStream();
            bmp.Save(ms, DrawingImageFormat.Png);
            ms.Position = 0;
            return Pix.LoadFromMemory(ms.ToArray());
        }

        private static Bitmap CropRightRegion(Bitmap src, double pct)
        {
            if (src == null) throw new ArgumentNullException(nameof(src));
            pct = Math.Clamp(pct, 0.05, 0.9);
            int w = src.Width;
            int h = src.Height;
            int x = Math.Max(0, (int)(w * (1.0 - pct)));
            int cw = w - x;
            var rect = new Rectangle(x, 0, cw, h);
            return src.Clone(rect, src.PixelFormat);
        }

        /// <summary>
        /// 高レベル処理: 画像全体で OCR を試行し、ParseByRules により抽出結果が得られなければ
        /// 右端を切り出して再試行するフォールバックを実施する。
        /// 戻り値: (ocrResult, parseMatches) - parseMatches が空ならマッチ無し。
        /// </summary>
        public async Task<(OcrResult ocrResult, List<ExtractMatch> parseMatches)> ProcessImageAndParseAsync(
            Bitmap originalBmp,
            OcrPreprocessor.Params prepParams,
            IEnumerable<ExtractRule> rules,
            CancellationToken ct = default,
            double cropRightPct = 0.32)
        {
            if (originalBmp == null) throw new ArgumentNullException(nameof(originalBmp));
            if (prepParams == null) prepParams = new OcrPreprocessor.Params();
            if (rules == null) rules = Enumerable.Empty<ExtractRule>();

            // Helper to run OCR on a preprocessed bitmap and return OcrResult + raw text
            (OcrResult, string) RunOcrOnBitmap(Bitmap bmp)
            {
                try
                {
                    using var pix = OcrPreprocessor.PreprocessToPix(bmp, prepParams);
                    using var page = _engine.Process(pix);
                    var result = new OcrResult
                    {
                        TimestampUtc = DateTime.UtcNow,
                        Region = new Rectangle(0, 0, bmp.Width, bmp.Height),
                        FullText = page.GetText()?.Trim() ?? ""
                    };

                    // populate words as in RecognizeAsync
                    using var iter = page.GetIterator();
                    iter.Begin();
                    do
                    {
                        string text = iter.GetText(PageIteratorLevel.Word);
                        if (string.IsNullOrWhiteSpace(text)) continue;

                        float confidence = iter.GetConfidence(PageIteratorLevel.Word);
                        if (iter.TryGetBoundingBox(PageIteratorLevel.Word, out var box))
                        {
                            var r = new Rectangle(box.X1, box.Y1, box.Width, box.Height);
                            result.Words.Add(new OcrWord
                            {
                                Text = text.Trim(),
                                Confidence = confidence,
                                BoundingBox = r
                            });
                        }
                        else
                        {
                            result.Words.Add(new OcrWord
                            {
                                Text = text.Trim(),
                                Confidence = confidence,
                                BoundingBox = new Rectangle(0, 0, 0, 0)
                            });
                        }
                    } while (iter.Next(PageIteratorLevel.Word));

                    return (result, result.FullText);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[OcrService] RunOcrOnBitmap failed: {ex}");
                    return (new OcrResult { TimestampUtc = DateTime.UtcNow, FullText = "" }, string.Empty);
                }
            }

            ct.ThrowIfCancellationRequested();

            // 1) Try full image
            try
            {
                using var preFull = OcrPreprocessor.Preprocess(originalBmp, prepParams);
                System.Diagnostics.Debug.WriteLine("[OcrService] Preprocessed full image");

                var (ocrFullResult, rawFull) = RunOcrOnBitmap(preFull);

                System.Diagnostics.Debug.WriteLine($"[OcrService] OCR Raw length={rawFull?.Length ?? 0} Raw:'{rawFull}'");

                if (!string.IsNullOrWhiteSpace(rawFull))
                {
                    var matchesFull = OcrParser.ParseByRules(rawFull, rules) ?? new List<ExtractMatch>();
                    System.Diagnostics.Debug.WriteLine($"[OcrService] Parse full results count={matchesFull.Count}");
                    if (matchesFull.Count > 0)
                    {
                        return (ocrFullResult, matchesFull);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OcrService] Full-image attempt error: {ex}");
            }

            ct.ThrowIfCancellationRequested();

            // 2) Fallback: crop right region and retry
            try
            {
                using var rightCrop = CropRightRegion(originalBmp, cropRightPct);
                System.Diagnostics.Debug.WriteLine("[OcrService] Created right-crop image for fallback");

                using var preCrop = OcrPreprocessor.Preprocess(rightCrop, prepParams);
                System.Diagnostics.Debug.WriteLine("[OcrService] Preprocessed cropped image");

                var (ocrCropResult, rawCrop) = RunOcrOnBitmap(preCrop);
                System.Diagnostics.Debug.WriteLine($"[OcrService] OCR Raw (crop) length={rawCrop?.Length ?? 0} Raw:'{rawCrop}'");

                if (!string.IsNullOrWhiteSpace(rawCrop))
                {
                    var matchesCrop = OcrParser.ParseByRules(rawCrop, rules) ?? new List<ExtractMatch>();
                    System.Diagnostics.Debug.WriteLine($"[OcrService] Parse crop results count={matchesCrop.Count}");
                    if (matchesCrop.Count > 0)
                    {
                        return (ocrCropResult, matchesCrop);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OcrService] Crop attempt error: {ex}");
            }

            // 3) none matched
            return (new OcrResult { TimestampUtc = DateTime.UtcNow, FullText = "" }, new List<ExtractMatch>());
        }

        public void Dispose()
        {
            if (_disposed) return;
            _engine?.Dispose();
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }
}