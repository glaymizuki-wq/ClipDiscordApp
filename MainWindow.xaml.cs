using ClipDiscordApp.Models;
using ClipDiscordApp.Parsers;
using ClipDiscordApp.Services;
using ClipDiscordApp.Utils;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Tesseract;
using YourNamespace.Ocr;

namespace ClipDiscordApp
{
    public partial class MainWindow : System.Windows.Window
    {
        // ---------- Fields ----------
        private readonly TesseractEngine _tesseractEngine;
        private readonly Dictionary<string, DateTime> _lastNotified = new();
        private readonly TimeSpan _notifyCooldown = TimeSpan.FromSeconds(5);
        private static readonly HttpClient _httpClient = new HttpClient();
        private CancellationTokenSource? _monitoringCts;
        private System.Drawing.Rectangle _selectedRegion = new System.Drawing.Rectangle(100, 100, 400, 100);
        private readonly string _rulesFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rules.json");
        private readonly string _logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ocr_logs");
        private MonitorService _monitorService;
        private CancellationTokenSource _monitorCts;
        private string _logsDir;
        private string? _lastPreviewHash = null;
        // ====== 追加フィールド（クラス内 Fields セクション付近に追記） ======
        private CancellationTokenSource? _previewCts;
        private Task? _previewTask;
        private readonly int _previewIntervalMs = 200; // UI 更新間隔（ms）
        private string? _previewLastHash = null;
        // 既にある場合は重複しないよう流用する
        private bool _skipFirstFrameAfterStart = true;    // 監視開始直後は判定をスキップするフラグ
        private bool _waitingForNextFrameAfterNotify = false; // 通知後は次の画像更新まで待つフラグ

        // Selection UI helpers
        private System.Windows.Point? _dragStart;
        private System.Windows.Shapes.Rectangle? _selectionRectVisual;

        // フィールド（クラス内）
        private string _lastSentMessage;
        private DateTime _lastSentAt = DateTime.MinValue;
        private readonly TimeSpan _duplicateWindow = TimeSpan.FromSeconds(10);
        
        private static readonly HttpClient _discordHttpClient = new HttpClient() { Timeout = TimeSpan.FromSeconds(8) };

        // ---------- Constructor ----------
        public MainWindow()
        {
            InitializeComponent();

            // diagnostic: ensure tessdata exists and log
            var exeDir = AppContext.BaseDirectory;
            Debug.WriteLine("ExeDir: " + exeDir);
            var tessDir = Path.Combine(exeDir, "tessdata");
            Debug.WriteLine("tessdata exists: " + Directory.Exists(tessDir));
            if (Directory.Exists(tessDir))
            {
                foreach (var f in Directory.GetFiles(tessDir))
                    Debug.WriteLine("tessdata file: " + f + " size=" + new FileInfo(f).Length);
            }

            // Ensure log dir
            Directory.CreateDirectory(_logDir);

            // Initialize Tesseract engine (use "jpn" if your tessdata contains jpn.traineddata)
            try
            {
                var tessDataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
                _tesseractEngine = new TesseractEngine(tessDataDir, "eng+jpn", EngineMode.Default);
                _tesseractEngine.SetVariable("tessedit_char_whitelist", "0123456789-:ABCDEFGHIJKLMNOPQRSTUVWXYZ ");
                _tesseractEngine.DefaultPageSegMode = PageSegMode.SingleLine;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Tesseract init failed: " + ex);
                throw;
            }
        }

        // ---------- Start / Stop Monitoring ----------
        public async Task StartMonitoringAsync(System.Drawing.Rectangle initialRegion, CancellationToken cancellationToken)
        {
            var rules = LoadRules();

            // Choose a fixed preprocessing preset for template creation and matching
            var prepParams = OcrPreprocessorPresets.ReadableText;

            Directory.CreateDirectory(_logDir);

            const double TEMPLATE_SCORE_THRESHOLD = 0.80;
            const double NOTIFY_THRESHOLD = 0.75;
            const double LOG_ONLY_THRESHOLD = 0.55;

            // ローカルヘルパ: candidate から Token/Label を安全に取り出す
            static string ReadCandidateToken(object cand)
            {
                if (cand == null) return string.Empty;
                var t = cand.GetType();
                string[] tokenNames = new[] { "Token", "token", "TokenValue", "Text", "Value" };
                foreach (var n in tokenNames)
                {
                    var p = t.GetProperty(n);
                    if (p != null && p.PropertyType == typeof(string)) return (string)(p.GetValue(cand) ?? string.Empty);
                    var f = t.GetField(n);
                    if (f != null && f.FieldType == typeof(string)) return (string)(f.GetValue(cand) ?? string.Empty);
                }
                var p2 = t.GetProperties().FirstOrDefault(pp => pp.PropertyType == typeof(string));
                if (p2 != null) return (string)(p2.GetValue(cand) ?? string.Empty);
                return string.Empty;
            }

            static string ReadCandidateLabel(object cand)
            {
                if (cand == null) return string.Empty;
                var t = cand.GetType();
                string[] labelNames = new[] { "Label", "label", "Target", "Kind" };
                foreach (var n in labelNames)
                {
                    var p = t.GetProperty(n);
                    if (p != null && p.PropertyType == typeof(string)) return (string)(p.GetValue(cand) ?? string.Empty);
                    var f = t.GetField(n);
                    if (f != null && f.FieldType == typeof(string)) return (string)(f.GetValue(cand) ?? string.Empty);
                }
                return string.Empty;
            }

            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                        try
                        {
#if !DEBUG
                            var tokyoNow = GetTokyoNow();
                            int hour = tokyoNow.Hour;
                            if (hour < 8 || hour >= 24)
                            {
                                // UI に一時的に表示（Dispatcher 経由）
                                try
                                {
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        StatusText.Text = $"Paused (time filter) {tokyoNow:HH:mm}";
                                    }));
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"[TimeFilter] Dispatcher update failed: {ex}");
                                }

                                // 1分ごとにチェック（キャンセル可能）
                                try
                                {
                                    await Task.Delay(TimeSpan.FromMinutes(1), cancellationToken);
                                }
                                catch (OperationCanceledException)
                                {
                                    // 呼び出し元のキャンセルをそのまま伝播（ループ外で適切に処理される想定）
                                    throw;
                                }
                                catch (Exception ex)
                                {
                                    // Delay 以外の一時的なエラーをログに残して継続
                                    Debug.WriteLine($"[TimeFilter] Delay failed: {ex}");
                                }

                                continue;
                            }
                            else
                            {
                                try
                                {
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        if (StatusText.Text.StartsWith("Paused (time filter)"))
                                            StatusText.Text = "Monitoring...";
                                    }));
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"[TimeFilter] Dispatcher update failed: {ex}");
                                }
                            }
                        }
                        catch (OperationCanceledException)
                        {
                            // 明示的にキャンセルは上位に伝えるかループを抜ける
                            throw;
                        }
                        catch (Exception ex)
                        {
                            // 想定外の例外はログに残して監視を継続
                            Debug.WriteLine($"[TimeFilter] unexpected error: {ex}");
                        }
#else
                            // Debug ビルド時はスキップ（ログを残す）
                            Debug.WriteLine("[TimeFilter] skipped in DEBUG build");
#endif

                            using var frame = CaptureFrameBitmap();
                        if (frame == null)
                        {
                            await Task.Delay(200, cancellationToken);
                            continue;
                        }

                        using var roi = CropToRegion(frame, initialRegion);
                        using var prepped = OcrPreprocessor.Preprocess(roi, prepParams);

                        // --- 画像変化トリガ（Insert here: after prepped is available） ---
                        string currentHash;
                        using (var tmpForHash = (Bitmap)prepped.Clone()) // ComputeSimpleHash expects Bitmap
                        {
                            currentHash = ComputeSimpleHash(tmpForHash);
                        }

                        // 起動直後は最初のフレームをキャプチャだけして判定を行わない
                        if (_skipFirstFrameAfterStart)
                        {
                            _lastPreviewHash = currentHash;
                            _skipFirstFrameAfterStart = false;
                            Debug.WriteLine("[Monitor] Warm-up: saved first frame hash, skipping detection");
                            await Task.Delay(200, cancellationToken);
                            continue;
                        }

                        // もし通知済みで、まだ次の画像更新を待っているならハッシュ変化で解除するまで待つ
                        if (_waitingForNextFrameAfterNotify)
                        {
                            if (_lastPreviewHash != currentHash)
                            {
                                Debug.WriteLine("[Monitor] Image changed after notify -> proceed");
                                _waitingForNextFrameAfterNotify = false;
                                _lastPreviewHash = currentHash;
                            }
                            else
                            {
                                // 画像変化なしなら判定スキップ
                                Debug.WriteLine("[Monitor] Waiting for next update after notify (no change)");
                                await Task.Delay(200, cancellationToken);
                                continue;
                            }
                        }
                        else
                        {
                            // 通常モード: 画像が同じならテンプレ/ OCR をスキップ
                            if (_lastPreviewHash != null && _lastPreviewHash == currentHash)
                            {
                                Debug.WriteLine("[Monitor] No visual change detected -> skip detection");
                                await Task.Delay(200, cancellationToken);
                                continue;
                            }

                            // 画像が変わっていれば次処理へ（かつ lastHash を更新）
                            _lastPreviewHash = currentHash;
                            Debug.WriteLine("[Monitor] Visual change detected -> run detection");
                            // Visual change detected -> update UI preview with the new preprocessed ROI image
                            try
                            {
                                // prepped is the Mat/Bitmap you already have for detection.
                                // If prepped is an OpenCvSharp Mat use appropriate conversion; here prepped is a Bitmap in your code path.
                                using var tmpForUi = (Bitmap)prepped.Clone();
                                var bs = BitmapToBitmapSource(tmpForUi); // 既存経路
                                if (bs.CanFreeze) bs.Freeze();

                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    try
                                    {
                                        CapturedImage.Source = bs;
                                        // OverlayCanvas サイズ同期が必要ならここで行う
                                        var src = PresentationSource.FromVisual(this);
                                        double dpiScaleX = src?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
                                        double dpiScaleY = src?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;
                                        OverlayCanvas.Width = Math.Round(bs.PixelWidth / dpiScaleX);
                                        OverlayCanvas.Height = Math.Round(bs.PixelHeight / dpiScaleY);
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"[UISet] failed: {ex}");
                                    }
                                }), DispatcherPriority.Render);
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[MonitorPreview] preview update failed: {ex}");
                            }

                            // optional: save debug image for visual inspection
                            try
                            {
                                var dbg = Path.Combine(_logDir, $"dbg_preview_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                                prepped.Save(dbg);
                                Debug.WriteLine($"[MonitorPreview] saved debug preview: {dbg}");
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[MonitorPreview] debug save failed: {ex.Message}");
                            }
                        }

                        // Debug save preprocessed
                        try
                        {
                            var fname = Path.Combine(_logDir, $"pre_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                            prepped.Save(fname);
                            Debug.WriteLine($"[OCR] Preprocessed image saved: {fname}");
                        }
                        catch (Exception ex) { Debug.WriteLine($"[OCR] Save prepped failed: {ex.Message}"); }

                        // ---------------- 1) Template match (priority) ----------------
                        try
                        {
                            TemplateMatcher.AcceptThreshold = TEMPLATE_SCORE_THRESHOLD;

                            int matchTimeoutSeconds = 4;
                            using var matchCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                            matchCts.CancelAfter(TimeSpan.FromSeconds(matchTimeoutSeconds));

                            var frameTag = DateTime.UtcNow.ToString("HHmmss_fff");
                            Debug.WriteLine($"[Template] Start match frame={frameTag} timeout={matchTimeoutSeconds}s");

                            MatchResult templateResult;
                            try
                            {
                                templateResult = await TemplateMatcher.CheckAsync(prepped, matchCts.Token);
                            }
                            catch (OperationCanceledException)
                            {
                                Debug.WriteLine($"[Template] matching canceled for frame={frameTag} (timeout or external). external={cancellationToken.IsCancellationRequested} matchCts={matchCts.IsCancellationRequested}");
                                await Task.Delay(100, CancellationToken.None);
                                continue;
                            }

                            if (templateResult != null && templateResult.Found)
                            {
                                Debug.WriteLine($"[Template] found {templateResult.Label} score={templateResult.BestScore:F3} tried={templateResult.TriedCandidates} elapsedMs={templateResult.ElapsedMs}");
                                var toParse = templateResult.Label.ToUpperInvariant();
                                var matches = OcrParser.ParseByRules(toParse, rules);
                                if (matches != null && matches.Any())
                                {
                                    foreach (var m in matches)
                                    {
                                        var key = !string.IsNullOrEmpty(m.RuleId) ? m.RuleId : m.RuleName;
                                        if (string.IsNullOrEmpty(key)) key = Guid.NewGuid().ToString();

                                        if (_lastNotified.TryGetValue(key, out var last) && (DateTime.UtcNow - last) < _notifyCooldown)
                                        {
                                            Debug.WriteLine($"[OCR] Skip notify for {key} (cooldown)");
                                            continue;
                                        }

                                        _lastNotified[key] = DateTime.UtcNow;
                                        _ = Task.Run(() => HandleMatchAsync(m));
                                        Debug.WriteLine($"[OCR] Match(from template): rule={m.RuleName} matches=[{string.Join(',', m.Matches)}]");
                                        _waitingForNextFrameAfterNotify = true;
                                        Debug.WriteLine("[Monitor] Notified -> will wait for next visual change before next detection");
                                    }
                                }

                                await Task.Delay(200, cancellationToken);
                                continue; // skip OCR for this frame
                            }
                            else
                            {
                                if (templateResult != null)
                                    Debug.WriteLine($"[Template] no confident template match (best {templateResult.Label} score={templateResult.BestScore:F3}), fallback to OCR");
                                else
                                    Debug.WriteLine($"[Template] templateResult null for frame={frameTag}, fallback to OCR");
                            }
                        }
                        catch (OperationCanceledException)
                        {
                            Debug.WriteLine("[Template] OperationCanceledException in template handling - continue monitoring");
                            await Task.Delay(100, CancellationToken.None);
                            continue;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[Template] matching error (non-fatal) {ex}");
                            await Task.Delay(100, cancellationToken);
                            continue;
                        }
                        // ---------------- end template matching ----------------

                        // ---------------- 2) OCR fallback using OcrTextUtils & OcrHelpers ----------------
                        try
                        {
                            // Get shared engine and convert bitmap to Pix
                            var engine = OcrTextUtils.GetEngine();
                            Pix pix = null;
                            try
                            {
                                pix = OcrTextUtils.BitmapToPix(prepped);
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[OCR] BitmapToPix failed: {ex}");
                                pix?.Dispose();
                                await Task.Delay(100, cancellationToken);
                                continue;
                            }

                            string ocrRaw;
                            try
                            {
                                ocrRaw = OcrHelpers.DoOcrWithRetries(engine, pix);
                            }
                            finally
                            {
                                pix.Dispose();
                            }

                            if (string.IsNullOrWhiteSpace(ocrRaw))
                            {
                                Debug.WriteLine("[OCR] DoOcrWithRetries returned empty");
                                await Task.Delay(100, cancellationToken);
                                continue;
                            }

                            Debug.WriteLine($"[OCR] Recognized raw text: {ocrRaw}");

                            var candidates = OcrHelpers.NormalizeAndDetectLabels(new[] { ocrRaw }, "ocr");
                            if (candidates == null || candidates.Count == 0)
                            {
                                Debug.WriteLine("[OCR] No label candidates from NormalizeAndDetectLabels");
                                await Task.Delay(100, cancellationToken);
                                continue;
                            }

                            // choose best candidate per label using OcrTextUtils.GetCandidateScore
                            var bestByLabel = candidates
                                .GroupBy(c => ReadCandidateLabel(c) ?? "")
                                .Select(g =>
                                {
                                    var best = g.OrderByDescending(x => OcrTextUtils.GetCandidateScore(x)).First();
                                    return new { Candidate = best, Score = OcrTextUtils.GetCandidateScore(best) };
                                })
                                .OrderByDescending(x => x.Score)
                                .Select(x => x.Candidate)
                                .ToList();

                            // log top candidates safely
                            foreach (var c in bestByLabel)
                            {
                                var tokenStr = ReadCandidateToken(c);
                                var labelStr = ReadCandidateLabel(c);
                                var scoreVal = OcrTextUtils.GetCandidateScore(c);
                                Debug.WriteLine($"[OCR] Candidate label={labelStr} token='{tokenStr}' score={scoreVal:F3} source=ocr");
                            }

                            // decide notify based on thresholds
                            object topCand = bestByLabel.FirstOrDefault();
                            double topScore = topCand != null ? OcrTextUtils.GetCandidateScore(topCand) : 0.0;
                            string topToken = topCand != null ? ReadCandidateToken(topCand) : string.Empty;

                            if (topScore >= NOTIFY_THRESHOLD && !string.IsNullOrWhiteSpace(topToken))
                            {
                                var toParse = topToken.ToUpperInvariant();
                                var matches = OcrParser.ParseByRules(toParse, rules);
                                if (matches != null && matches.Any())
                                {
                                    foreach (var m in matches)
                                    {
                                        var key = !string.IsNullOrEmpty(m.RuleId) ? m.RuleId : m.RuleName;
                                        if (string.IsNullOrEmpty(key)) key = Guid.NewGuid().ToString();

                                        if (_lastNotified.TryGetValue(key, out var last) && (DateTime.UtcNow - last) < _notifyCooldown)
                                        {
                                            Debug.WriteLine($"[OCR] Skip notify for {key} (cooldown)");
                                            continue;
                                        }

                                        _lastNotified[key] = DateTime.UtcNow;
                                        _ = Task.Run(() => HandleMatchAsync(m));
                                        Debug.WriteLine($"[OCR] Match(from OCR): rule={m.RuleName} matches=[{string.Join(',', m.Matches)}]");
                                        _waitingForNextFrameAfterNotify = true;
                                        Debug.WriteLine("[Monitor] Notified -> will wait for next visual change before next detection");
                                    }
                                }
                            }
                            else
                            {
                                Debug.WriteLine($"[OCR] Low-confidence top candidate (score={topScore:F3}) - log only");
                            }

                            await Task.Delay(200, cancellationToken);
                        }
                        catch (OperationCanceledException)
                        {
                            Debug.WriteLine("[OCR] OCR operation canceled for this frame - continue monitoring");
                            await Task.Delay(100, CancellationToken.None);
                            continue;
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[OCR] OCR processing failed (non-fatal): {ex}");
                            await Task.Delay(100, cancellationToken);
                            continue;
                        }
                        // ---------------- end OCR fallback ----------------
                    }
                    catch (OperationCanceledException)
                    {
                        Debug.WriteLine("[StartMonitoring] External cancellation requested - stopping monitoring.");
                        break;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[StartMonitoring] Non-fatal exception in loop: {ex}");
                        await Task.Delay(200, cancellationToken);
                        continue;
                    }
                } // end whil
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine("[StartMonitoring] Monitoring canceled (top-level).");
            }
            finally
            {
                Debug.WriteLine("[StartMonitoring] Exiting StartMonitoringAsync.");
            }
        }

        // シンプルな差分ハッシュ（高速で概ねの変化検出用）
        private string ComputeSimpleHash(Bitmap bmp)
        {
            int sampleStepsX = Math.Max(1, bmp.Width / 32);
            int sampleStepsY = Math.Max(1, bmp.Height / 32);
            unchecked
            {
                int h = 17;
                for (int y = 0; y < bmp.Height; y += sampleStepsY)
                {
                    for (int x = 0; x < bmp.Width; x += sampleStepsX)
                    {
                        var p = bmp.GetPixel(Math.Min(x, bmp.Width - 1), Math.Min(y, bmp.Height - 1));
                        h = h * 31 + p.R;
                        h = h * 31 + p.G;
                        h = h * 31 + p.B;
                    }
                }
                return h.ToString("X8");
            }
        }

        // ---------- Utilities / Helpers ----------
        private IEnumerable<ExtractRule> LoadRules()
        {
            try
            {
                if (File.Exists(_rulesFileName))
                {
                    var json = File.ReadAllText(_rulesFileName, Encoding.UTF8);
                    var list = JsonSerializer.Deserialize<List<ExtractRule>>(json);
                    if (list != null && list.Any()) return list;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LoadRules] Failed to load rules.json: {ex.Message}");
            }

            return new List<ExtractRule>
            {
                new ExtractRule { Id = "r1", Name = "Sell", Pattern = "SELL", Type = ExtractRuleType.Keyword, Enabled = true, Order = 0 },
                new ExtractRule { Id = "r2", Name = "Buy", Pattern = "BUY", Type = ExtractRuleType.Keyword, Enabled = true, Order = 1 }
            };
        }

        private Bitmap CaptureFrameBitmap()
        {
            var bounds = Screen.PrimaryScreen.Bounds;
            var bmp = new Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(bmp);
            g.CopyFromScreen(bounds.X, bounds.Y, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy);
            return bmp;
        }

        private Bitmap CropToRegion(Bitmap src, System.Drawing.Rectangle region)
        {
            if (region.Width > 0 && region.Height > 0)
            {
                var bmp = new Bitmap(region.Width, region.Height, src.PixelFormat);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.DrawImage(src, 0, 0, region, GraphicsUnit.Pixel);
                }
                return bmp;
            }
            else
            {
                return new Bitmap(src);
            }
        }

        // 置き換え用メソッド
        private async Task HandleMatchAsync(ExtractMatch match)
        {
            try
            {
                Debug.WriteLine($"[HandleMatchAsync] Handling match rule={match.RuleName} text={string.Join(",", match.Matches)}");

                // マッチ文字列の決定: BUY/SELL 優先で探す
                string matchText = null;
                if (match?.Matches != null && match.Matches.Count > 0)
                {
                    matchText = match.Matches.FirstOrDefault(m => string.Equals(m, "BUY", StringComparison.OrdinalIgnoreCase))
                                ?? match.Matches.FirstOrDefault(m => string.Equals(m, "SELL", StringComparison.OrdinalIgnoreCase))
                                ?? match.Matches.FirstOrDefault();
                }
                if (string.IsNullOrWhiteSpace(matchText))
                {
                    Debug.WriteLine("[HandleMatchAsync] no match text to send");
                    return;
                }

                // 重複抑止
                if (_lastSentMessage == matchText && (DateTime.UtcNow - _lastSentAt) < _duplicateWindow)
                {
                    Debug.WriteLine("[HandleMatchAsync] suppressed duplicate message");
                    return;
                }
                _lastSentMessage = matchText;
                _lastSentAt = DateTime.UtcNow;

                // Webhook を取得して送信
                var webhookUrl = GetWebhookUrl("DetectionChannel");
                if (string.IsNullOrWhiteSpace(webhookUrl))
                {
                    Debug.WriteLine("[HandleMatchAsync] DetectionChannel webhook not configured");
                    return;
                }

                using var cts = CancellationTokenSource.CreateLinkedTokenSource(_monitoringCts?.Token ?? CancellationToken.None);
                cts.CancelAfter(TimeSpan.FromSeconds(15));

                try
                {
                    var ok = await NotifyDiscordAsync(webhookUrl, matchText, cts.Token);
                    Debug.WriteLine($"[HandleMatchAsync] discord notify result={ok}");
                }
                catch (OperationCanceledException) { Debug.WriteLine("[HandleMatchAsync] notify cancelled"); }
                catch (Exception ex) { Debug.WriteLine($"[HandleMatchAsync] notify error: {ex}"); }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[HandleMatchAsync] Error: {ex}");
            }
        }

        private BitmapSource BitmapToBitmapSource(Bitmap bmp)
        {
            var handle = bmp.GetHbitmap();
            try
            {
                return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    handle, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            }
            finally
            {
                NativeMethods.DeleteObject(handle);
            }
        }

        // ---------- XAML Event Handlers (UI binding) ----------
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            InitSelectionVisual();
            _logsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            Directory.CreateDirectory(_logsDir);

            var templatesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            TemplateMatcher.Initialize(templatesPath);

            BtnStartMonitoring.IsEnabled = true;
            BtnStopMonitoring.IsEnabled = false;
            StatusText.Text = "Ready";

            try
            {
                var src = PresentationSource.FromVisual(this);
                double dpiScaleX = 1.0, dpiScaleY = 1.0;
                if (src?.CompositionTarget != null)
                {
                    dpiScaleX = src.CompositionTarget.TransformToDevice.M11;
                    dpiScaleY = src.CompositionTarget.TransformToDevice.M22;
                }
                Debug.WriteLine($"[Startup] DPI scaleX={dpiScaleX:F2} scaleY={dpiScaleY:F2}");

                Debug.WriteLine($"[Startup] initial _selectedRegion: {_selectedRegion}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Startup] DPI debug failed: {ex.Message}");
            }
        }

        private void BtnStopMonitoring_Click(object sender, RoutedEventArgs e)
        {
            // 1) 停止対象のプレビュータスクを確実に止める
            StopPreviewLoop();

            // 2) 監視ループをキャンセル
            try { _monitoringCts?.Cancel(); } catch { }

            // 3) 内部フラグを必要ならリセット（次回開始時の安全策）
            _skipFirstFrameAfterStart = false;
            _waitingForNextFrameAfterNotify = false;
            _lastPreviewHash = null;
            _previewLastHash = null;

            // 4) UI 更新
            BtnStartMonitoring.IsEnabled = true;
            BtnStopMonitoring.IsEnabled = false;
            StatusText.Text = "Stopped";

            Debug.WriteLine("[BtnStop] Monitoring and preview stopped");
        }

        private void BtnSelectRegion_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selector = new RegionSelectorWindow { Owner = this };
                var dlg = selector.ShowDialog();
                if (dlg == true)
                {
                    var sel = ConvertDrawingRectToRectangle(selector.SelectedRegion);
                    _selectedRegion = sel;
                    StatusText.Text = $"Region selected: {_selectedRegion.X},{_selectedRegion.Y} {_selectedRegion.Width}x{_selectedRegion.Height}";
                    DrawSelectionOnOverlay(_selectedRegion);
                    // Immediate preview of the selected region (show one snapshot)
                    try
                    {
                        using var full = CaptureFrameBitmap();
                        using var crop = CropToRegion(full, _selectedRegion);
                        var bs = BitmapToBitmapSource(crop);
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            try { CapturedImage.Source = bs; }
                            catch (Exception ex) { Debug.WriteLine($"[StartPreview] UI set failed: {ex.Message}"); }
                        }), DispatcherPriority.Render);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[StartPreview] capture/display failed: {ex}");
                    }
                }
                else
                {
                    StatusText.Text = "Region selection cancelled";
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[BtnSelectRegion] Error: {ex}");
                System.Windows.MessageBox.Show($"領域選択に失敗しました: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnTestNotification_Click(object sender, RoutedEventArgs e)
        {
            BtnTestNotification.IsEnabled = false;
            try
            {
                var webhookUrl = GetWebhookUrl("TestChannel");
                if (string.IsNullOrWhiteSpace(webhookUrl))
                {
                    System.Windows.MessageBox.Show("Webhook URL が設定されていません。", "設定エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var userText = (txtNotificationContent?.Text ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(userText))
                {
                    System.Windows.MessageBox.Show("送信するメッセージを入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    txtNotificationContent.Focus();
                    return;
                }

                var now = DateTime.Now;
                var contentText = $"{now:yyyy-MM-dd HH:mm:ss} {userText}";

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                bool ok = false;

                try
                {
                    ok = await NotifyDiscordAsync(webhookUrl, contentText, cts.Token);
                }
                catch (OperationCanceledException)
                {
                    Debug.WriteLine("[BtnTestNotification] Discord send cancelled/timed out");
                    System.Windows.MessageBox.Show("送信がキャンセルまたはタイムアウトしました。", "送信中断", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[BtnTestNotification] notify exception: {ex}");
                    System.Windows.MessageBox.Show($"例外が発生しました: {ex.Message}", "例外", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (ok)
                {
                    Debug.WriteLine("[BtnTestNotification] Discord送信成功");
                    System.Windows.MessageBox.Show("送信完了", "通知", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    Debug.WriteLine("[BtnTestNotification] Discord送信失敗（詳細はログを確認してください）");
                    System.Windows.MessageBox.Show("送信失敗（詳細はログを確認してください）", "送信エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            finally
            {
                BtnTestNotification.IsEnabled = true;
            }
        }

        private void BtnOpenRules_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (File.Exists(_rulesFileName))
                {
                    Process.Start(new ProcessStartInfo("notepad.exe", _rulesFileName) { UseShellExecute = true });
                }
                else
                {
                    System.Windows.MessageBox.Show("rules.json が見つかりません。", "ルールファイル", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[BtnOpenRules] Error: {ex}");
            }
        }

        // ---------- Image mouse handlers for region selection (preview) ----------
        private void CapturedImage_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var img = sender as System.Windows.Controls.Image;
            _dragStart = e.GetPosition(img);

            if (_selectionRectVisual == null)
            {
                _selectionRectVisual = new System.Windows.Shapes.Rectangle
                {
                    Stroke = System.Windows.Media.Brushes.Red,
                    StrokeThickness = 2,
                    StrokeDashArray = new System.Windows.Media.DoubleCollection { 4, 2 },
                    Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 255, 0, 0))
                };
                OverlayCanvas.Children.Add(_selectionRectVisual);
            }

            img?.CaptureMouse();
        }

        private void CapturedImage_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            var img = sender as System.Windows.Controls.Image;
            if (!_dragStart.HasValue || _selectionRectVisual == null) return;
            var pos = e.GetPosition(img);

            var x = Math.Min(_dragStart.Value.X, pos.X);
            var y = Math.Min(_dragStart.Value.Y, pos.Y);
            var w = Math.Abs(_dragStart.Value.X - pos.X);
            var h = Math.Abs(_dragStart.Value.Y - pos.Y);

            System.Windows.Controls.Canvas.SetLeft(_selectionRectVisual, x);
            System.Windows.Controls.Canvas.SetTop(_selectionRectVisual, y);
            _selectionRectVisual.Width = w;
            _selectionRectVisual.Height = h;
        }

        private void CapturedImage_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var img = sender as System.Windows.Controls.Image;
            if (!_dragStart.HasValue || _selectionRectVisual == null)
            {
                _dragStart = null;
                img?.ReleaseMouseCapture();
                return;
            }

            var pos = e.GetPosition(img);
            var x = (int)Math.Min(_dragStart.Value.X, pos.X);
            var y = (int)Math.Min(_dragStart.Value.Y, pos.Y);
            var w = (int)Math.Abs(_dragStart.Value.X - pos.X);
            var h = (int)Math.Abs(_dragStart.Value.Y - pos.Y);

            if (w > 4 && h > 4)
            {
                // Note: CapturedImage preview must be 1:1 for screen coords to match. If scaled, conversion required.
                _selectedRegion = new System.Drawing.Rectangle(x, y, w, h);
                StatusText.Text = $"Region selected: {_selectedRegion.X},{_selectedRegion.Y} {_selectedRegion.Width}x{_selectedRegion.Height}";
                DrawSelectionOnOverlay(_selectedRegion);
            }
            else
            {
                StatusText.Text = "Region selection cancelled (too small)";
                OverlayCanvas.Children.Remove(_selectionRectVisual);
                _selectionRectVisual = null;
            }

            _dragStart = null;
            img?.ReleaseMouseCapture();
        }

        // ---------- Helpers ----------
        private System.Drawing.Rectangle ConvertDrawingRectToRectangle(DrawingRect dr)
        {
            return new System.Drawing.Rectangle(dr.X, dr.Y, dr.Width, dr.Height);
        }

        private void DrawSelectionOnOverlay(System.Drawing.Rectangle rect)
        {
            if (_selectionRectVisual != null)
            {
                OverlayCanvas.Children.Remove(_selectionRectVisual);
                _selectionRectVisual = null;
            }

            var r = new System.Windows.Shapes.Rectangle
            {
                Stroke = System.Windows.Media.Brushes.Lime,
                StrokeThickness = 2,
                Fill = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 0, 255, 0))
            };

            OverlayCanvas.Children.Add(r);
            System.Windows.Controls.Canvas.SetLeft(r, rect.X);
            System.Windows.Controls.Canvas.SetTop(r, rect.Y);
            r.Width = rect.Width;
            r.Height = rect.Height;
            _selectionRectVisual = r;
        }

        // ---------- Native helper ----------
        private static class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("gdi32.dll")]
            public static extern bool DeleteObject(IntPtr hObject);
        }

        private string GetWebhookUrl(string key = "ClipDiscordApp")
        {
            try
            {
                // 探索候補順: 実行フォルダ -> %APPDATA%\<App>\ -> 環境変数
                var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                var candidates = new[]
                {
            Path.Combine(exeDir, "discord_webhooks.json"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ClipDiscordApp", "discord_webhooks.json")
        };

                foreach (var configPath in candidates)
                {
                    if (!File.Exists(configPath))
                    {
                        Debug.WriteLine($"[GetWebhookUrl] config not found: {configPath}");
                        continue;
                    }

                    string json;
                    try
                    {
                        json = File.ReadAllText(configPath, Encoding.UTF8);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[GetWebhookUrl] could not read {configPath}: {ex}");
                        continue;
                    }

                    try
                    {
                        using var doc = JsonDocument.Parse(json);
                        if (doc.RootElement.TryGetProperty(key, out var val) && val.ValueKind == JsonValueKind.String)
                        {
                            var url = val.GetString()?.Trim();
                            if (!string.IsNullOrEmpty(url) && Uri.TryCreate(url, UriKind.Absolute, out var u) && (u.Scheme == "https" || u.Scheme == "http"))
                            {
                                Debug.WriteLine($"[GetWebhookUrl] loaded webhook from {configPath}");
                                return url;
                            }
                            Debug.WriteLine($"[GetWebhookUrl] invalid or empty url for key '{key}' in {configPath}");
                        }
                        else
                        {
                            Debug.WriteLine($"[GetWebhookUrl] key '{key}' not found or not a string in {configPath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[GetWebhookUrl] parse error in {configPath}: {ex}");
                    }
                }

                // 環境変数フォールバック
                var envKey = $"DISCORD_WEBHOOK_{key.ToUpperInvariant()}";
                var envUrl = Environment.GetEnvironmentVariable(envKey);
                if (!string.IsNullOrWhiteSpace(envUrl) && Uri.TryCreate(envUrl.Trim(), UriKind.Absolute, out var envUri))
                {
                    Debug.WriteLine($"[GetWebhookUrl] loaded webhook from ENV {envKey}");
                    return envUrl.Trim();
                }

                Debug.WriteLine("[GetWebhookUrl] webhook not found in any candidate locations");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetWebhookUrl] unexpected error: {ex}");
            }

            return null;
        }

        private void BtnStartMonitoring_Click(object sender, RoutedEventArgs e)
        {
            StartPreviewLoop();
            _skipFirstFrameAfterStart = true;
            _waitingForNextFrameAfterNotify = false;
            _lastPreviewHash = null;

            var selector = new RegionSelectorWindow { Owner = this };
            var dlg = selector.ShowDialog();
            if (dlg != true)
            {
                StatusText.Text = "Monitoring cancelled (no region selected)";
                return;
            }

            var d = selector.SelectedRegion;
            var selRect = new System.Windows.Rect(d.X, d.Y, d.Width, d.Height);
            _selectedRegion = ConvertImageRectToBitmapRect(selRect, CapturedImage);
            DrawSelectionOnOverlay(_selectedRegion);

            // --- immediate snapshot of the selected region and set overlay size ---
            try
            {
                using var full = CaptureFrameBitmap();
                using var crop = CropToRegion(full, _selectedRegion);
                var bs = BitmapToBitmapSource(crop); // 既存経路
                if (bs.CanFreeze) bs.Freeze();

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        CapturedImage.Source = bs;
                        // OverlayCanvas サイズ同期が必要ならここで行う
                        var src = PresentationSource.FromVisual(this);
                        double dpiScaleX = src?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
                        double dpiScaleY = src?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;
                        OverlayCanvas.Width = Math.Round(bs.PixelWidth / dpiScaleX);
                        OverlayCanvas.Height = Math.Round(bs.PixelHeight / dpiScaleY);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[UISet] failed: {ex}");
                    }
                }), DispatcherPriority.Render);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[StartPreview] capture/display failed outer: {ex.Message}");
            }

            BtnStartMonitoring.IsEnabled = false;
            BtnStopMonitoring.IsEnabled = true;
            StatusText.Text = "Monitoring...";

            try { _monitoringCts?.Cancel(); } catch { }

            _monitoringCts = new CancellationTokenSource();

            if (_monitorService == null)
            {
                _monitorService = new MonitorService(
                    _logsDir,
                    _tesseractEngine,
                    () => CaptureFrameBitmap(),
                    (bmp, rect) => CropToRegion(bmp, rect),
                    (bmp) => ComputeSimpleHash(bmp),
                    (bmp) => BitmapToBitmapSource(bmp),
                    async (object m) => { await HandleMatchAsync((ExtractMatch)m); },
                    () => LoadRules(),
                    (bmpSource) =>
                    {
                        // Preview callback from MonitorService: update UI and overlay size
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            try
                            {
                                CapturedImage.Source = bmpSource;

                                var src = PresentationSource.FromVisual(this);
                                double dpiScaleX = src?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
                                double dpiScaleY = src?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;

                                OverlayCanvas.Width = Math.Round(bmpSource.PixelWidth / dpiScaleX);
                                OverlayCanvas.Height = Math.Round(bmpSource.PixelHeight / dpiScaleY);

                                Debug.WriteLine($"[MainWindow] preview set {bmpSource.PixelWidth}x{bmpSource.PixelHeight}, overlay {OverlayCanvas.Width}x{OverlayCanvas.Height}");
                            }
                            catch (Exception ex) { Debug.WriteLine($"[MainWindow] preview set failed: {ex.Message}"); }
                        }), DispatcherPriority.Render);
                    });
            }

            _ = Task.Run(async () =>
            {
                try
                {
                    await StartMonitoringAsync(_selectedRegion, _monitoringCts.Token);
                }
                catch (OperationCanceledException)
                {
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[BtnStart] Error: {ex}");
                    Dispatcher.Invoke(() => StatusText.Text = "Error");
                }
                finally
                {
                    Dispatcher.Invoke(() =>
                    {
                        BtnStartMonitoring.IsEnabled = true;
                        BtnStopMonitoring.IsEnabled = false;
                        if (!_monitoringCts.IsCancellationRequested) StatusText.Text = "Stopped";
                    });
                }
            });
        }

        // 追加メソッド：Image上の選択矩形を元画像（BitmapSource）ピクセル座標に変換して返す
        private System.Drawing.Rectangle ConvertImageRectToBitmapRect(System.Windows.Rect selRectInImage, System.Windows.Controls.Image imageControl)
        {
            if (imageControl?.Source is System.Windows.Media.Imaging.BitmapSource bmp)
            {
                double displayW = imageControl.ActualWidth;
                double displayH = imageControl.ActualHeight;

                double sourceW = bmp.PixelWidth;
                double sourceH = bmp.PixelHeight;

                var src = PresentationSource.FromVisual(this);
                double dpiScaleX = 1.0, dpiScaleY = 1.0;
                if (src?.CompositionTarget != null)
                {
                    dpiScaleX = src.CompositionTarget.TransformToDevice.M11;
                    dpiScaleY = src.CompositionTarget.TransformToDevice.M22;
                }

                double displayPxW = displayW * dpiScaleX;
                double displayPxH = displayH * dpiScaleY;

                double scaleX = sourceW / Math.Max(1.0, displayPxW);
                double scaleY = sourceH / Math.Max(1.0, displayPxH);

                int sx = (int)Math.Round(selRectInImage.X * dpiScaleX * scaleX);
                int sy = (int)Math.Round(selRectInImage.Y * dpiScaleY * scaleY);
                int sw = (int)Math.Round(selRectInImage.Width * dpiScaleX * scaleX);
                int sh = (int)Math.Round(selRectInImage.Height * dpiScaleY * scaleY);

                return new System.Drawing.Rectangle(
                    Math.Max(0, sx),
                    Math.Max(0, sy),
                    Math.Max(1, sw),
                    Math.Max(1, sh)
                );
            }

            return new System.Drawing.Rectangle(
                (int)Math.Max(0, selRectInImage.X),
                (int)Math.Max(0, selRectInImage.Y),
                (int)Math.Max(1, selRectInImage.Width),
                (int)Math.Max(1, selRectInImage.Height)
            );
        }

        // ====== InitSelectionVisual の追加（Window_Loaded から呼ぶ） ======
        private void InitSelectionVisual()
        {
            // OverlayCanvas に選択矩形用のビジュアルがまだなければ準備する
            try
            {
                // 既に _selectionRectVisual を使っているため、ここでは何もしない場合が多いが
                // OverlayCanvas の初期状態確認や背景設定などを行う余地を残す。
                OverlayCanvas.Background = System.Windows.Media.Brushes.Transparent;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[InitSelectionVisual] {ex}");
            }
        }

        // ====== プレビュー開始/停止 ======
        private void StartPreviewLoop()
        {
            if (_previewTask != null && !_previewTask.IsCompleted) return;
            _previewCts = new CancellationTokenSource();
            var ct = _previewCts.Token;
            _previewTask = Task.Run(async () =>
            {
                while (!ct.IsCancellationRequested)
                {
                    try
                    {
                        using var bmp = CaptureFrameBitmap();
                        if (bmp != null)
                        {
                            // If a selected region exists, crop; otherwise show full screen
                            Bitmap toShow;
                            if (_selectedRegion.Width > 0 && _selectedRegion.Height > 0)
                            {
                                var safeR = new System.Drawing.Rectangle(
                                    Math.Max(0, Math.Min(_selectedRegion.X, bmp.Width - 1)),
                                    Math.Max(0, Math.Min(_selectedRegion.Y, bmp.Height - 1)),
                                    Math.Max(1, Math.Min(_selectedRegion.Width, Math.Max(1, bmp.Width - _selectedRegion.X))),
                                    Math.Max(1, Math.Min(_selectedRegion.Height, Math.Max(1, bmp.Height - _selectedRegion.Y)))
                                );
                                try
                                {
                                    toShow = bmp.Clone(safeR, bmp.PixelFormat);
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"[Preview] crop failed, using full bmp: {ex.Message}");
                                    toShow = (Bitmap)bmp.Clone();
                                }
                            }
                            else
                            {
                                toShow = (Bitmap)bmp.Clone();
                            }

                            // 変化検出（軽いハッシュ）: 変わっていなければ UI 更新をスキップ
                            string hash = ComputeSimpleHash(toShow);
                            if (_previewLastHash != null && _previewLastHash == hash)
                            {
                                Debug.WriteLine("[Preview] hash unchanged -> skip UI update");
                                toShow.Dispose();
                            }
                            else
                            {
                                _previewLastHash = hash;
                                var bs = BitmapToBitmapSource(toShow); // 既存経路
                                if (bs.CanFreeze) bs.Freeze();

                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    try
                                    {
                                        CapturedImage.Source = bs;
                                        // OverlayCanvas サイズ同期が必要ならここで行う
                                        var src = PresentationSource.FromVisual(this);
                                        double dpiScaleX = src?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
                                        double dpiScaleY = src?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;
                                        OverlayCanvas.Width = Math.Round(bs.PixelWidth / dpiScaleX);
                                        OverlayCanvas.Height = Math.Round(bs.PixelHeight / dpiScaleY);
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"[UISet] failed: {ex}");
                                    }
                                }), DispatcherPriority.Render);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[PreviewLoop] Exception: {ex}");
                    }

                    try { await Task.Delay(_previewIntervalMs, ct); }
                    catch (TaskCanceledException) { break; }
                }
            }, ct);
        }

        private void StopPreviewLoop()
        {
            try
            {
                if (_previewCts != null && !_previewCts.IsCancellationRequested)
                {
                    _previewCts.Cancel();
                    try { _previewTask?.Wait(300); } catch { }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[StopPreviewLoop] {ex}");
            }
            finally
            {
                _previewTask = null;
                _previewCts?.Dispose();
                _previewCts = null;
            }
        }

        private async Task<bool> NotifyDiscordAsync(string webhookUrl, string content, CancellationToken ct, int maxRetries = 3)
        {
            if (string.IsNullOrWhiteSpace(webhookUrl)) { Debug.WriteLine("[Discord] webhook empty"); return false; }

            var payload = new { content = content };
            var json = System.Text.Json.JsonSerializer.Serialize(payload);
            var attempt = 0;
            var backoff = TimeSpan.FromSeconds(1);

            while (attempt < maxRetries)
            {
                attempt++;
                using var body = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
                try
                {
                    using var resp = await _discordHttpClient.PostAsync(webhookUrl, body, ct);
                    if (resp.IsSuccessStatusCode) return true;

                    if ((int)resp.StatusCode == 401)
                    {
                        Debug.WriteLine("[Discord] Unauthorized 401 - webhook invalid");
                        return false;
                    }

                    if ((int)resp.StatusCode == 429)
                    {
                        if (resp.Headers.RetryAfter?.Delta is TimeSpan delta) await Task.Delay(delta, ct);
                        else await Task.Delay(backoff, ct);
                        backoff = backoff * 2;
                        continue;
                    }

                    if ((int)resp.StatusCode >= 400 && (int)resp.StatusCode < 500)
                    {
                        var txt = await resp.Content.ReadAsStringAsync(ct);
                        Debug.WriteLine($"[Discord] client error {(int)resp.StatusCode}: {txt}");
                        return false;
                    }

                    var serverBody = await resp.Content.ReadAsStringAsync(ct);
                    Debug.WriteLine($"[Discord] server error {(int)resp.StatusCode}: {serverBody}");
                }
                catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
                catch (Exception ex) { Debug.WriteLine($"[Discord] send exception: {ex}"); }

                try { await Task.Delay(backoff, ct); } catch (OperationCanceledException) { throw; }
                backoff = backoff * 2;
            }

            Debug.WriteLine("[Discord] send failed after retries");
            return false;
        }

        private DateTime GetTokyoNow()
        {
            try
            {
                // Windows/Linux で ID が異なるため両方試す
                TimeZoneInfo tz = null;
                try { tz = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time"); } catch { }
                if (tz == null)
                {
                    try { tz = TimeZoneInfo.FindSystemTimeZoneById("Asia/Tokyo"); } catch { }
                }
                if (tz != null) return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, tz);
            }
            catch { }
            // フェールセーフでローカル時刻（開発環境向け）
            return DateTime.Now;
        }
    }
}