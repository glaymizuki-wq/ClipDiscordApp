using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Tesseract;
using ClipDiscordApp.Models;
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
        }

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

        public void Dispose()
        {
            if (_disposed) return;
            _engine?.Dispose();
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }
}