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
    // デリゲート型群: UI 側の実装を渡すため
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

        // 外部（MainWindow）から渡す処理
        private readonly CaptureFrameBitmapDelegate _captureFrame;
        private readonly CropToRegionDelegate _cropToRegion;
        private readonly ComputeSimpleHashDelegate _computeHash;
        private readonly BitmapToBitmapSourceDelegate _bmpToSource;
        private readonly HandleMatchAsyncDelegate _handleMatchAsync;
        private readonly LoadRulesDelegate _loadRules;

        // UI 更新用アクション（Dispatcher 内で実行される）
        private readonly Action<BitmapSource> _setPreviewAction;

        // コンストラクタ: 必要な外部機能を注入する（MainWindow から渡す）
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

            // Tesseract whitelist 設定（推奨）: 初期化時に一度設定
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

        // StartMonitoringAsync 本体（先に提示したロジックをそのまま使用）
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

                    // prepped 保存
                    string preppedPath = null;
                    try
                    {
                        preppedPath = Path.Combine(_logDir, $"pre_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                        prepped.Save(preppedPath);
                        System.Diagnostics.Debug.WriteLine($"[OCR] Preprocessed image saved: {preppedPath}");
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[OCR] Save prepped failed: {ex.Message}"); }

                    // gray / bin も保存
                    string grayPath = null, binPath = null;
                    try
                    {
                        using var grayBmp = OcrPreprocessorExtensions.GetGrayBitmap(prepped);
                        if (grayBmp != null)
                        {
                            grayPath = Path.Combine(_logDir, $"gray_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                            grayBmp.Save(grayPath);
                            System.Diagnostics.Debug.WriteLine($"[OCR] Gray image saved: {grayPath}");
                        }
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[OCR] Save gray failed: {ex.Message}"); }

                    try
                    {
                        using var binBmp = OcrPreprocessorExtensions.GetBinaryBitmap(prepped);
                        if (binBmp != null)
                        {
                            binPath = Path.Combine(_logDir, $"bin_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                            binBmp.Save(binPath);
                            System.Diagnostics.Debug.WriteLine($"[OCR] Binary image saved: {binPath}");
                        }
                    }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[OCR] Save bin failed: {ex.Message}"); }

                    // 1) TemplateMatcher 先行
                    // 置き換え用（非同期・タイムアウト付き）
                    try
                    {
                        // TemplateMatcher の閾値を反映（任意）
                        TemplateMatcher.AcceptThreshold = TEMPLATE_SCORE_THRESHOLD;

                        // この単一マッチのための短期タイムアウトを作る
                        using var matchCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                        matchCts.CancelAfter(TimeSpan.FromSeconds(2)); // 必要に応じて調整

                        System.Diagnostics.Debug.WriteLine($"[Template] Start check (timeout=2s)");
                        var templateRes = await TemplateMatcher.CheckAsync(prepped, matchCts.Token);

                        System.Diagnostics.Debug.WriteLine($"[Template] result Found={templateRes.Found} Label={templateRes.Label} BestScore={templateRes.BestScore:F3} Tried={templateRes.TriedCandidates} ElapsedMs={templateRes.ElapsedMs}");

                        if (templateRes.Found && templateRes.BestScore >= TEMPLATE_SCORE_THRESHOLD)
                        {
                            var toParse = templateRes.Label;
                            var matches = OcrParser.ParseByRules(toParse, rules);
                            if (matches != null && matches.Any())
                            {
                                foreach (var m in matches)
                                {
                                    var key = !string.IsNullOrEmpty(m.RuleId) ? m.RuleId : m.RuleName;
                                    if (string.IsNullOrEmpty(key)) key = Guid.NewGuid().ToString();

                                    if (_lastNotified.TryGetValue(key, out DateTime last) && (DateTime.UtcNow - last) < _notifyCooldown)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[OCR] Skip notify for {key} (cooldown)");
                                        continue;
                                    }

                                    _lastNotified[key] = DateTime.UtcNow;
                                    _ = _handleMatchAsync(m);
                                    System.Diagnostics.Debug.WriteLine($"[OCR] Match(from template): rule={m.RuleName} matches=[{string.Join(',', m.Matches)}]");
                                }
                            }

                            await Task.Delay(200, cancellationToken);
                            continue;
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        // タイムアウトまたは外部キャンセル。ログだけ出して次フレームへ（必要ならリトライを入れる）
                        System.Diagnostics.Debug.WriteLine("[Template] check canceled (timeout or external cancellation)");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[Template] check failed: {ex}");
                    }

                    // 2) フォールバック OCR: binary と gray を試す
                    string textBin = string.Empty;
                    string textGray = string.Empty;

                    try
                    {
                        using var pixBin = PixConverter.ToPix(prepped);
                        textBin = OcrHelpers.DoOcrWithRetries(_tesseractEngine, pixBin) ?? string.Empty;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OCR] binary OCR failed: {ex.Message}");
                    }

                    try
                    {
                        using var grayBmp = OcrPreprocessorExtensions.GetGrayBitmap(prepped);
                        using var pixGray = PixConverter.ToPix(grayBmp);
                        textGray = OcrHelpers.DoOcrWithRetries(_tesseractEngine, pixGray) ?? string.Empty;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OCR] gray OCR failed: {ex.Message}");
                    }

                    System.Diagnostics.Debug.WriteLine($"[OCR] Raw bin len={textBin.Length} bin:'{(textBin.Length > 200 ? textBin.Substring(0, 200) + "..." : textBin)}'");
                    System.Diagnostics.Debug.WriteLine($"[OCR] Raw gray len={textGray.Length} gray:'{(textGray.Length > 200 ? textGray.Substring(0, 200) + "..." : textGray)}'");

                    // 3) Normalize & Label scoring
                    var rawCandidates = new List<(string text, string source)>();
                    rawCandidates.Add((textBin, "binary"));
                    rawCandidates.Add((textGray, "gray"));

                    var labelCandidates = new List<LabelCandidate>();
                    foreach (var (txt, src) in rawCandidates)
                    {
                        if (string.IsNullOrWhiteSpace(txt)) continue;
                        try
                        {
                            var list = OcrHelpers.NormalizeAndDetectLabels(new[] { txt }, src);
                            if (list != null) labelCandidates.AddRange(list);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[OcrHelpers] NormalizeAndDetectLabels failed for {src}: {ex.Message}");
                        }
                    }

                    System.Diagnostics.Debug.WriteLine("[OCR] LabelCandidates (top 10):");
                    foreach (var c in labelCandidates.OrderByDescending(x => x.Confidence).Take(10))
                        System.Diagnostics.Debug.WriteLine($"  -> {c}");

                    var bestSell = labelCandidates.Where(c => c.Label == "SELL").OrderByDescending(c => c.Confidence).FirstOrDefault();
                    var bestBuy = labelCandidates.Where(c => c.Label == "BUY").OrderByDescending(c => c.Confidence).FirstOrDefault();
                    var chosen = (bestSell?.Confidence ?? 0) >= (bestBuy?.Confidence ?? 0) ? bestSell : bestBuy;

                    System.Diagnostics.Debug.WriteLine($"[OCR] bestSell={bestSell?.Text}:{bestSell?.Confidence:F3} (src={bestSell?.Source})");
                    System.Diagnostics.Debug.WriteLine($"[OCR] bestBuy ={bestBuy?.Text}:{bestBuy?.Confidence:F3} (src={bestBuy?.Source})");
                    System.Diagnostics.Debug.WriteLine($"[OCR] chosen   ={chosen?.Label}:{chosen?.Text}:{chosen?.Confidence:F3}");

                    if (chosen != null && chosen.Confidence >= NOTIFY_THRESHOLD)
                    {
                        var toParse = chosen.Text;
                        var matches = OcrParser.ParseByRules(toParse, rules);
                        if (matches != null && matches.Any())
                        {
                            foreach (var m in matches)
                            {
                                var key = !string.IsNullOrEmpty(m.RuleId) ? m.RuleId : m.RuleName;
                                if (string.IsNullOrEmpty(key)) key = Guid.NewGuid().ToString();
                                if (_lastNotified.TryGetValue(key, out DateTime last) && (DateTime.UtcNow - last) < _notifyCooldown)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[OCR] Skip notify for {key} (cooldown)");
                                    continue;
                                }
                                _lastNotified[key] = DateTime.UtcNow;
                                _ = _handleMatchAsync(m);
                                System.Diagnostics.Debug.WriteLine($"[OCR] Match: rule={m.RuleName} matches=[{string.Join(',', m.Matches)}]");
                            }
                        }
                    }
                    else if (chosen != null && chosen.Confidence >= LOG_ONLY_THRESHOLD)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OCR] Low-confidence candidate (log only): {chosen.Label} '{chosen.Text}' conf={chosen.Confidence:F3}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[OCR] No confident candidate found");
                    }

                    // プレビュー更新（UI スレッドで渡す）
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
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[Preview] BitmapToBitmapSource failed: {ex.Message}");
                                    bmpSource = null;
                                }
                                if (bmpSource != null)
                                {
                                    // UI スレッドで preview を設定する
                                    try { _setPreviewAction(bmpSource); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[Preview] set action failed: {ex.Message}"); }
                                }
                            }
                        }
                        finally
                        {
                            try { bmpClone?.Dispose(); } catch { }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[Preview] Error computing/updating preview: {ex}");
                    }

                    await Task.Delay(200, cancellationToken);
                }
                catch (TaskCanceledException tex)
                {
                    System.Diagnostics.Debug.WriteLine($"[StartMonitoring] TaskCanceledException: {tex.Message}\n{tex.StackTrace}");
                    break;
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[StartMonitoring] Error: {ex}");
                    await Task.Delay(500, cancellationToken);
                }
            }
        }
    }
}