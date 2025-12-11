
using ClipDiscordApp.Models;
using ClipDiscordApp.Parsers;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Tesseract;
using YourNamespace.Ocr;

namespace ClipDiscordApp.Services
{
    public delegate Bitmap CaptureFrameBitmapDelegate();
    public delegate Bitmap CropToRegionDelegate(Bitmap src, Rectangle region);
    public delegate string ComputeSimpleHashDelegate(Bitmap bmp);
    public delegate BitmapSource BitmapToBitmapSourceDelegate(Bitmap bmp);
    public delegate Task HandleMatchAsyncDelegate(object match);
    public delegate dynamic LoadRulesDelegate();

    public class MonitorService
    {
        private readonly string _logDir;
        private readonly TesseractEngine _tesseractEngine;
        private readonly TimeSpan _notifyCooldown = TimeSpan.FromSeconds(30);
        private readonly Dictionary<string, DateTime> _lastNotified = new();
        private string _lastPreviewHash = string.Empty;

        private readonly CaptureFrameBitmapDelegate _captureFrame;
        private readonly CropToRegionDelegate _cropToRegion;
        private readonly ComputeSimpleHashDelegate _computeHash;
        private readonly BitmapToBitmapSourceDelegate _bmpToSource;
        private readonly HandleMatchAsyncDelegate _handleMatchAsync;
        private readonly LoadRulesDelegate _loadRules;
        private readonly Action<BitmapSource> _setPreviewAction;

        public MonitorService(
            string logDir,
            TesseractEngine tessEngine,
            CaptureFrameBitmapDelegate captureFrame,
            CropToRegionDelegate cropToRegion,
            ComputeSimpleHashDelegate computeHash,
            BitmapToBitmapSourceDelegate bmpToSource,
            HandleMatchAsyncDelegate handleMatchAsync,
            LoadRulesDelegate loadRules,
            Action<BitmapSource> setPreviewAction)
        {
            _logDir = logDir ?? throw new ArgumentNullException(nameof(logDir));
            _tesseractEngine = tessEngine ?? throw new ArgumentNullException(nameof(tessEngine));
            _captureFrame = captureFrame ?? throw new ArgumentNullException(nameof(captureFrame));
            _cropToRegion = cropToRegion ?? throw new ArgumentNullException(nameof(cropToRegion));
            _computeHash = computeHash ?? throw new ArgumentNullException(nameof(computeHash));
            _bmpToSource = bmpToSource ?? throw new ArgumentNullException(nameof(bmpToSource));
            _handleMatchAsync = handleMatchAsync ?? throw new ArgumentNullException(nameof(handleMatchAsync));
            _loadRules = loadRules ?? throw new ArgumentNullException(nameof(loadRules));
            _setPreviewAction = setPreviewAction ?? throw new ArgumentNullException(nameof(setPreviewAction));

            try
            {
                _tesseractEngine.SetVariable("tessedit_char_whitelist", "0123456789-:ABCDEFGHIJKLMNOPQRSTUVWXYZ ");
                _tesseractEngine.SetVariable("classify_bln_numeric_mode", "0");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MonitorService] tess variable set failed: {ex.Message}");
            }
        }

        public async Task StartMonitoringAsync(Rectangle initialRegion, CancellationToken cancellationToken)
        {
            var rules = _loadRules();

            var prepParams = new OcrPreprocessor.Params
            {
                Scale = 5,
                AdaptiveBlockSize = 11,
                AdaptiveC = 5.0,
                UseClahe = false,
                MorphKernel = 2,
                UseDilate = false,
                DilateIterations = 0,
                Border = 8,
                Sharpen = true,
                SavePreprocessed = true,
                PreprocessedFilenamePrefix = "pre_morph3_v2"
            };

            Directory.CreateDirectory(_logDir);

            const double TEMPLATE_SCORE_THRESHOLD = 0.75;
            const double NOTIFY_THRESHOLD = 0.75;
            const double LOG_ONLY_THRESHOLD = 0.55;

            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    using var frame = _captureFrame();
                    if (frame == null)
                    {
                        await Task.Delay(200, cancellationToken);
                        continue;
                    }

                    using var roi = _cropToRegion(frame, initialRegion);
                    using var prepped = OcrPreprocessor.Preprocess(roi, prepParams);

#if DEBUG
                    // DEBUGビルド時のみ画像保存
                    try
                    {
                        string preppedPath = Path.Combine(_logDir, $"pre_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                        prepped.Save(preppedPath);
                        System.Diagnostics.Debug.WriteLine($"[OCR] Preprocessed image saved: {preppedPath}");

                        using var grayBmp = OcrPreprocessorExtensions.GetGrayBitmap(prepped);
                        if (grayBmp != null)
                        {
                            string grayPath = Path.Combine(_logDir, $"gray_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                            grayBmp.Save(grayPath);
                            System.Diagnostics.Debug.WriteLine($"[OCR] Gray image saved: {grayPath}");
                        }

                        using var binBmp = OcrPreprocessorExtensions.GetBinaryBitmap(prepped);
                        if (binBmp != null)
                        {
                            string binPath = Path.Combine(_logDir, $"bin_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                            binBmp.Save(binPath);
                            System.Diagnostics.Debug.WriteLine($"[OCR] Binary image saved: {binPath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OCR] Save debug images failed: {ex.Message}");
                    }
#endif

                    // TemplateMatcher
                    try
                    {
                        TemplateMatcher.AcceptThreshold = TEMPLATE_SCORE_THRESHOLD;
                        using var matchCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                        matchCts.CancelAfter(TimeSpan.FromSeconds(2));

                        var templateRes = await TemplateMatcher.CheckAsync(prepped, matchCts.Token);
                        if (templateRes.Found && templateRes.BestScore >= TEMPLATE_SCORE_THRESHOLD)
                        {
                            var matches = OcrParser.ParseByRules(templateRes.Label, rules);
                            if (matches != null && matches.Any())
                            {
                                foreach (var m in matches)
                                {
                                    var key = !string.IsNullOrEmpty(m.RuleId) ? m.RuleId : m.RuleName;
                                    if (string.IsNullOrEmpty(key)) key = Guid.NewGuid().ToString();

                                    if (_lastNotified.TryGetValue(key, out DateTime last) && (DateTime.UtcNow - last) < _notifyCooldown)
                                        continue;

                                    _lastNotified[key] = DateTime.UtcNow;
                                    _ = _handleMatchAsync(m);
                                }
                            }
                            await Task.Delay(200, cancellationToken);
                            continue;
                        }
                    }
                    catch (OperationCanceledException) { }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[Template] check failed: {ex}"); }

                    // OCR fallback
                    string textBin = string.Empty;
                    string textGray = string.Empty;

                    try
                    {
                        using var pixBin = PixConverter.ToPix(prepped);
                        textBin = OcrHelpers.DoOcrWithRetries(_tesseractEngine, pixBin) ?? string.Empty;
                    }
                    catch { }

                    try
                    {
                        using var grayBmp = OcrPreprocessorExtensions.GetGrayBitmap(prepped);
                        using var pixGray = PixConverter.ToPix(grayBmp);
                        textGray = OcrHelpers.DoOcrWithRetries(_tesseractEngine, pixGray) ?? string.Empty;
                    }
                    catch { }

                    var rawCandidates = new List<(string text, string source)> { (textBin, "binary"), (textGray, "gray") };
                    var labelCandidates = new List<LabelCandidate>();
                    foreach (var (txt, src) in rawCandidates)
                    {
                        if (!string.IsNullOrWhiteSpace(txt))
                        {
                            var list = OcrHelpers.NormalizeAndDetectLabels(new[] { txt }, src);
                            if (list != null) labelCandidates.AddRange(list);
                        }
                    }

                    var bestSell = labelCandidates.Where(c => c.Label == "SELL").OrderByDescending(c => c.Confidence).FirstOrDefault();
                    var bestBuy = labelCandidates.Where(c => c.Label == "BUY").OrderByDescending(c => c.Confidence).FirstOrDefault();
                    var chosen = (bestSell?.Confidence ?? 0) >= (bestBuy?.Confidence ?? 0) ? bestSell : bestBuy;

                    if (chosen != null && chosen.Confidence >= NOTIFY_THRESHOLD)
                    {
                        var matches = OcrParser.ParseByRules(chosen.Text, rules);
                        if (matches != null && matches.Any())
                        {
                            foreach (var m in matches)
                            {
                                var key = !string.IsNullOrEmpty(m.RuleId) ? m.RuleId : m.RuleName;
                                if (string.IsNullOrEmpty(key)) key = Guid.NewGuid().ToString();
                                if (_lastNotified.TryGetValue(key, out DateTime last) && (DateTime.UtcNow - last) < _notifyCooldown)
                                    continue;
                                _lastNotified[key] = DateTime.UtcNow;
                                _ = _handleMatchAsync(m);
                            }
                        }
                    }

                    // プレビュー更新
                    try
                    {
                        Bitmap bmpClone = null;
                        try
                        {
                            bmpClone = (Bitmap)roi.Clone();
                            var hash = _computeHash(bmpClone);
                            if (hash != _lastPreviewHash)
                            {
                                _lastPreviewHash = hash;
                                BitmapSource bmpSource = null;
                                try
                                {
                                    bmpSource = _bmpToSource(bmpClone);
                                    if (bmpSource != null && bmpSource.CanFreeze) bmpSource.Freeze();
                                }
                                catch { bmpSource = null; }
                                if (bmpSource != null)
                                {
                                    try { _setPreviewAction(bmpSource); } catch { }
                                }
                            }
                        }
                        finally { bmpClone?.Dispose(); }
                    }
                    catch { }

                    await Task.Delay(200, cancellationToken);
                }
                catch (TaskCanceledException) { break; }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[StartMonitoring] Error: {ex}");
                    await Task.Delay(500, cancellationToken);
                }
            }
        }
    }
}
