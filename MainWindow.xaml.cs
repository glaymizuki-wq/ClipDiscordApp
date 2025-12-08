using ClipDiscordApp.Models;
using ClipDiscordApp.Parsers;
using ClipDiscordApp.Services;
using ClipDiscordApp.Utils;
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
        private readonly System.Collections.Concurrent.ConcurrentDictionary<string, DateTime> _lastNotified = new();
        private readonly TimeSpan _notifyCooldown = TimeSpan.FromSeconds(5);
        private static readonly HttpClient _httpClient = new HttpClient();
        private CancellationTokenSource? _monitoringCts;
        private System.Drawing.Rectangle _selectedRegion = new System.Drawing.Rectangle(100, 100, 400, 100);
        private readonly string _rulesFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rules.json");

        // 保存先をアプリベースの Logs に統一
        private readonly string _logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

        private MonitorService _monitorService;
        private CancellationTokenSource _monitorCts;
        private string? _lastPreviewHash = null;

        // ====== 追加フィールド ======
        private CancellationTokenSource? _previewCts;
        private Task? _previewTask;
        private readonly int _previewIntervalMs = 200; // UI 更新間隔（ms）
        private string? _previewLastHash = null;
        private bool _skipFirstFrameAfterStart = true;    // 監視開始直後は判定スキップ
        private bool _waitingForNextFrameAfterNotify = false; // 通知後は次の画像更新まで待つ
                                                              // 追加フィールド（MainWindow クラス内）
        private readonly object _previewLock = new object();
        private DateTime _lastPreviewUpdate = DateTime.MinValue;
        private readonly TimeSpan _minPreviewInterval = TimeSpan.FromMilliseconds(100);

        // Selection UI helpers
        private System.Windows.Point? _dragStart;
        private System.Windows.Shapes.Rectangle? _selectionRectVisual;

        // 重複送信防止
        private string _lastSentMessage;
        private DateTime _lastSentAt = DateTime.MinValue;
        private readonly TimeSpan _duplicateWindow = TimeSpan.FromSeconds(10);

        private bool _cleanupScheduled = false;
        private readonly object _cleanupLock = new object();
        private volatile bool _allowMonitorPreview = false;

        private static readonly HttpClient _discordHttpClient = new HttpClient() { Timeout = TimeSpan.FromSeconds(8) };

        private string? _monitorLastHash; // 監視用ハッシュ（StartMonitoringAsync が更新）

        // 画像保持日数（基本1日）
        private int _debugKeepDays = 1;

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

            // 保存先ディレクトリの用意（Logs 統一）
            Directory.CreateDirectory(_logDir);

            // Tesseract 初期化（eng+jpn）
            try
            {
                var tessDataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
                _tesseractEngine = new TesseractEngine(tessDataDir, "eng+jpn", EngineMode.Default);

                // 日本語ラベルをOCRしたい場合、ホワイトリストは外す方が安全
                // _tesseractEngine.SetVariable("tessedit_char_whitelist", "0123456789-:ABCDEFGHIJKLMNOPQRSTUVWXYZ ");

                _tesseractEngine.DefaultPageSegMode = PageSegMode.SingleLine;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Tesseract init failed: " + ex);
                throw;
            }

            try
            {
                lock (_cleanupLock)
                {
                    if (!_cleanupScheduled)
                    {
                        CleanupOldDebugFiles(_logDir, _debugKeepDays);
                        StartDailyCleanupTask();
                        _cleanupScheduled = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[InitCleanup] failed: {ex}");
            }
        }

        // ---------- Start / Stop Monitoring ----------
        // StartMonitoringAsync - 監視ループ（roi をハッシュ元にし、デバッグログを追加）
        public async Task StartMonitoringAsync(System.Drawing.Rectangle initialRegion, CancellationToken cancellationToken)
        {
            var rules = LoadRules();
            var prepParams = OcrPreprocessorPresets.ReadableText;
            Directory.CreateDirectory(_logDir);

            const double TEMPLATE_SCORE_THRESHOLD = 0.80;

            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
#if !DEBUG
            var tokyoNow = GetTokyoNow();
            int hour = tokyoNow.Hour;
            if (hour < 9 || hour >= 24)
            {
                Dispatcher.BeginInvoke(new Action(() => StatusText.Text = $"Paused (time filter) {tokyoNow:HH:mm}"));
                await Task.Delay(TimeSpan.FromMinutes(1), cancellationToken);
                continue;
            }
            else
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (StatusText.Text.StartsWith("Paused (time filter)"))
                        StatusText.Text = "Monitoring...";
                }));
            }
#else
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

                    // ハッシュは roi（生画像）を元に計算する
                    string currentHash;
                    try
                    {
                        using var tmpForHash = (Bitmap)roi.Clone();
                        currentHash = ComputeSimpleHash(tmpForHash);
                        Debug.WriteLine($"[MonitorImageInfo] roiSize={tmpForHash.Width}x{tmpForHash.Height} checksum={ComputeQuickChecksum(tmpForHash)} hash={currentHash}");
                        Debug.WriteLine($"[PreprocessInfo] preppedSize={prepped.Width}x{prepped.Height}");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[Monitor] hash compute failed: {ex}");
                        await Task.Delay(200, cancellationToken);
                        continue;
                    }

                    Debug.WriteLine($"[DebugHashes] preview={_previewLastHash} monitor={_monitorLastHash} computed={currentHash} waiting={_waitingForNextFrameAfterNotify}");

                    if (_skipFirstFrameAfterStart)
                    {
                        _monitorLastHash = currentHash;
                        _skipFirstFrameAfterStart = false;
                        Debug.WriteLine("[Monitor] Warm-up: saved first frame hash, skipping detection");
                        await Task.Delay(200, cancellationToken);
                        continue;
                    }

                    if (_waitingForNextFrameAfterNotify)
                    {
                        if (_monitorLastHash != currentHash)
                        {
                            Debug.WriteLine("[Monitor] Image changed after notify -> proceed");
                            _waitingForNextFrameAfterNotify = false;
                            _monitorLastHash = currentHash;
                        }
                        else
                        {
                            Debug.WriteLine("[Monitor] Waiting for next update after notify (no change)");
                            await Task.Delay(200, cancellationToken);
                            continue;
                        }
                    }
                    else
                    {
                        if (_monitorLastHash != null && _monitorLastHash == currentHash)
                        {
                            Debug.WriteLine("[Monitor] No visual change detected -> skip detection");
                            await Task.Delay(200, cancellationToken);
                            continue;
                        }

                        _monitorLastHash = currentHash;
                        Debug.WriteLine("[Monitor] Visual change detected -> run detection");
                    }

                    // UI プレビュー更新（prepped を表示）
                    try
                    {
                        using var tmpForUi = (Bitmap)prepped.Clone();
                        var bs = BitmapToBitmapSource(tmpForUi);
                        if (bs != null && bs.CanFreeze) bs.Freeze();
                        UpdatePreview(bs);
                    }
                    catch (Exception ex) { Debug.WriteLine($"[MonitorPreview] preview update failed: {ex}"); }

                    // テンプレートマッチ優先
                    try
                    {
                        TemplateMatcher.AcceptThreshold = TEMPLATE_SCORE_THRESHOLD;
                        using var matchCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                        matchCts.CancelAfter(TimeSpan.FromSeconds(4));

                        var frameTag = DateTime.Now.ToString("HHmmss_fff");
                        Debug.WriteLine($"[Template] Start match frame={frameTag}");

                        MatchResult templateResult = null;
                        try { templateResult = await TemplateMatcher.CheckAsync(prepped, matchCts.Token); }
                        catch (OperationCanceledException)
                        {
                            Debug.WriteLine($"[Template] matching canceled for frame={frameTag}");
                            await Task.Delay(100, CancellationToken.None);
                            continue;
                        }

                        if (templateResult != null && templateResult.Found)
                        {
                            Debug.WriteLine($"[Template] found {templateResult.Label} score={templateResult.BestScore:F3}");
                            var toParse = templateResult.Label;
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
                                    Debug.WriteLine($"[OCR] Match(from template): rule={m.RuleName} matches=[{string.Join(',', m.Matches)}]");
                                    _ = Task.Run(async () => await HandleMatchAsync(m));
                                    _waitingForNextFrameAfterNotify = true;
                                    Debug.WriteLine("[Monitor] Notified -> will wait for next visual change before next detection");
                                }
                            }

                            await Task.Delay(200, cancellationToken);
                            continue;
                        }
                    }
                    catch (Exception ex) { Debug.WriteLine($"[Template] match error: {ex}"); }

                    // OCR フォールバック
                    try
                    {
                        using var tmpForTess = (Bitmap)prepped.Clone();
                        using var pix = PixConverter.ToPix(tmpForTess);
                        using var page = _tesseractEngine.Process(pix);
                        var rawText = page.GetText()?.Trim() ?? string.Empty;

                        var matches = OcrParser.ParseByRules(rawText, rules);
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
                                Debug.WriteLine($"[OCR] Match(from OCR): rule={m.RuleName} matches=[{string.Join(',', m.Matches)}]");
                                _ = Task.Run(async () => await HandleMatchAsync(m));
                                _waitingForNextFrameAfterNotify = true;
                                Debug.WriteLine("[Monitor] Notified -> will wait for next visual change before next detection");
                            }
                        }
                    }
                    catch (Exception ex) { Debug.WriteLine($"[OCR] error: {ex}"); }

                    await Task.Delay(200, cancellationToken);
                }
            }
            catch (OperationCanceledException) { Debug.WriteLine("[Monitor] canceled"); }
            catch (Exception ex) { Debug.WriteLine($"[Monitor] unexpected: {ex}"); }
        }

        // 領域情報を含めるオーバーロード（推奨）
        // 領域情報を含める実装（新しい名前）
        private string ComputeSimpleHashWithRegion(Bitmap bmp, System.Drawing.Rectangle region)
        {
            unchecked
            {
                int h = 17;
                h = h * 23 + region.X;
                h = h * 23 + region.Y;
                h = h * 23 + region.Width;
                h = h * 23 + region.Height;

                int sampleStepsX = Math.Max(1, bmp.Width / 32);
                int sampleStepsY = Math.Max(1, bmp.Height / 32);
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

        // 既存互換用（既存呼び出しを壊さない）
        // ComputeSimpleHash - 差分ハッシュ dHash 実装（感度向上のため）
        private string ComputeSimpleHash(System.Drawing.Bitmap bmp)
        {
            if (bmp == null) return string.Empty;

            const int w = 9;
            const int h = 8;

            using var resized = new System.Drawing.Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            using (var g = System.Drawing.Graphics.FromImage(resized))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBilinear;
                g.DrawImage(bmp, 0, 0, w, h);
            }

            var bits = new System.Text.StringBuilder();
            for (int y = 0; y < h; y++)
            {
                for (int x = 0; x < w - 1; x++)
                {
                    var c1 = resized.GetPixel(x, y);
                    var c2 = resized.GetPixel(x + 1, y);
                    int v1 = (c1.R + c1.G + c1.B) / 3;
                    int v2 = (c2.R + c2.G + c2.B) / 3;
                    bits.Append(v1 > v2 ? '1' : '0');
                }
            }

            var bitStr = bits.ToString();
            var sb = new System.Text.StringBuilder(bitStr.Length / 4 + 1);
            for (int i = 0; i < bitStr.Length; i += 4)
            {
                string nibble = bitStr.Substring(i, Math.Min(4, bitStr.Length - i)).PadRight(4, '0');
                sb.Append(Convert.ToInt32(nibble, 2).ToString("X1"));
            }
            return sb.ToString();
        }

        // ComputeQuickChecksum - デバッグ用簡易チェックサム
        private string ComputeQuickChecksum(System.Drawing.Bitmap bmp)
        {
            if (bmp == null) return "NULL";
            long sum = 0;
            int stepY = Math.Max(1, bmp.Height / 10);
            int stepX = Math.Max(1, bmp.Width / 10);

            try
            {
                for (int y = 0; y < bmp.Height; y += stepY)
                {
                    for (int x = 0; x < bmp.Width; x += stepX)
                    {
                        var c = bmp.GetPixel(x, y);
                        sum += c.R + c.G + c.B;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ComputeQuickChecksum] error: {ex}");
                return "ERR";
            }

            return (sum & 0xFFFFFFFF).ToString("X8");
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

            // デフォルトは日本語ラベル（上昇中／下落中）
            return new List<ExtractRule>
            {
                new ExtractRule { Id = "up",   Name = "上昇中", Pattern = "上昇中", Type = ExtractRuleType.Keyword, Enabled = true, Order = 0 },
                new ExtractRule { Id = "down", Name = "下落中", Pattern = "下落中", Type = ExtractRuleType.Keyword, Enabled = true, Order = 1 }
            };
        }

        private Bitmap CaptureFrameBitmap()
        {
            var bounds = Screen.PrimaryScreen.Bounds;
            using var raw = new Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(raw))
            {
                g.CopyFromScreen(bounds.X, bounds.Y, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy);
            }

            var clone = new Bitmap(raw.Width, raw.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var g2 = Graphics.FromImage(clone))
            {
                g2.DrawImage(raw, 0, 0, raw.Width, raw.Height);
            }
            return clone;
        }

        private Bitmap CropToRegion(Bitmap src, System.Drawing.Rectangle region)
        {
            if (region.Width > 0 && region.Height > 0)
            {
                var bmp = new Bitmap(region.Width, region.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.DrawImage(src, new System.Drawing.Rectangle(0, 0, region.Width, region.Height), region, GraphicsUnit.Pixel);
                }
                return bmp;
            }
            else
            {
                var clone = new Bitmap(src.Width, src.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(clone))
                {
                    g.DrawImage(src, 0, 0, src.Width, src.Height);
                }
                return clone;
            }
        }

        // MatchResultItem を受け取る本体
        private async Task HandleMatchAsync(MatchResultItem m)
        {
            // --- Matches 補完 ---
            if (m.RuleId == "sell" && (m.Matches == null || m.Matches.Length == 0))
            {
                m.Matches = new[] { "SELL" };
            }
            if (m.RuleId == "buy" && (m.Matches == null || m.Matches.Length == 0))
            {
                m.Matches = new[] { "BUY" };
            }

            System.Diagnostics.Debug.WriteLine(
                $"[HandleMatch] start rule={m?.RuleName} ruleId={m?.RuleId} matches={string.Join(",", m?.Matches ?? new string[0])}");

            try
            {
                var ok = await Notifier.SendOrderAsync(m).ConfigureAwait(false);
                System.Diagnostics.Debug.WriteLine($"[HandleMatch] finished rule={m?.RuleName} success={ok}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[HandleMatch] Exception: {ex}");
            }
        }

        // ExtractMatch を受け取るオーバーロード（こちらはそのまま）
        private async Task HandleMatchAsync(object extractMatch)
        {
            var item = ToMatchResultItem(extractMatch);
            if (item == null)
            {
                System.Diagnostics.Debug.WriteLine("[HandleMatch] conversion from ExtractMatch failed");
                return;
            }
            await HandleMatchAsync(item).ConfigureAwait(false);
        }

        private BitmapSource BitmapToBitmapSource(Bitmap bmp)
        {
            var handle = bmp.GetHbitmap();
            try
            {
                var bs = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    handle, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                if (bs.CanFreeze) bs.Freeze();
                return bs;
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
            Directory.CreateDirectory(_logDir);

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

            // 起動時クリーンアップと日次タスク開始
            try
            {
                CleanupOldDebugFiles(_logDir, _debugKeepDays);
                StartDailyCleanupTask();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Startup] Cleanup scheduling failed: {ex}");
            }

            // 起動時に Webhook を初期化してログ出力
            try
            {
                Notifier.InitializeWebhook("DetectionChannel");
                Debug.WriteLine($"[Init] Notifier.WebhookUrl={(string.IsNullOrWhiteSpace(Notifier.WebhookUrl) ? "EMPTY" : Notifier.WebhookUrl)}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Init] Notifier.InitializeWebhook failed: {ex}");
            }
        }

        private void BtnStopMonitoring_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1) プレビュー停止（StartPreviewLoop を止める）
                try
                {
                    _previewCts?.Cancel();
                    _previewCts?.Dispose();
                    _previewCts = null;
                    Debug.WriteLine("[PreviewCancel] preview cancelled");
                }
                catch (ObjectDisposedException ex)
                {
                    Debug.WriteLine($"[PreviewCancel] ObjectDisposedException: {ex}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[PreviewCancel] unexpected error: {ex}");
                }

                // 2) StartPreviewLoop を止めたので MonitorService のプレビューを許可
                _allowMonitorPreview = true;

                // 3) 監視ループも停止するならここで
                try
                {
                    _monitoringCts?.Cancel();
                    _monitoringCts?.Dispose();
                    _monitoringCts = null;
                    Debug.WriteLine("[MonitorCancel] monitoring cancelled");
                }
                catch (ObjectDisposedException ex)
                {
                    Debug.WriteLine($"[MonitorCancel] ObjectDisposedException: {ex}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[MonitorCancel] unexpected error: {ex}");
                }

                // UI 更新
                BtnStartMonitoring.IsEnabled = true;
                BtnStopMonitoring.IsEnabled = false;
                StatusText.Text = "Stopped";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[BtnStop] outer error: {ex}");
            }
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

                    // ハッシュとフラグをリセットして、選択直後の古いハッシュ判定を防ぐ
                    _previewLastHash = null;
                    _lastPreviewHash = null;
                    _skipFirstFrameAfterStart = true;
                    _waitingForNextFrameAfterNotify = false;

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
        // 追加メソッド：Image上の選択矩形を元画像（BitmapSource）ピクセル座標に変換して返す
        private System.Drawing.Rectangle ConvertImageRectToBitmapRect(System.Windows.Rect selRectInImage, System.Windows.Controls.Image imageControl)
        {
            try
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

                    double displayPxW = Math.Max(1.0, displayW * dpiScaleX);
                    double displayPxH = Math.Max(1.0, displayH * dpiScaleY);

                    double scaleX = sourceW / displayPxW;
                    double scaleY = sourceH / displayPxH;

                    int sx = (int)Math.Round(selRectInImage.X * dpiScaleX * scaleX);
                    int sy = (int)Math.Round(selRectInImage.Y * dpiScaleY * scaleY);
                    int sw = (int)Math.Round(selRectInImage.Width * dpiScaleX * scaleX);
                    int sh = (int)Math.Round(selRectInImage.Height * dpiScaleY * scaleY);

                    sx = Math.Max(0, sx);
                    sy = Math.Max(0, sy);
                    sw = Math.Max(1, sw);
                    sh = Math.Max(1, sh);

                    if (sx + sw > sourceW) sw = (int)Math.Max(1, sourceW - sx);
                    if (sy + sh > sourceH) sh = (int)Math.Max(1, sourceH - sy);

                    return new System.Drawing.Rectangle(sx, sy, sw, sh);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ConvertImageRectToBitmapRect] {ex}");
            }

            return new System.Drawing.Rectangle(
                (int)Math.Max(0, selRectInImage.X),
                (int)Math.Max(0, selRectInImage.Y),
                (int)Math.Max(1, selRectInImage.Width),
                (int)Math.Max(1, selRectInImage.Height)
            );
        }

        // 初期化
        private void InitSelectionVisual()
        {
            try
            {
                OverlayCanvas.Background = System.Windows.Media.Brushes.Transparent;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[InitSelectionVisual] {ex}");
            }
        }

        // ====== プレビュー開始/停止 ======
        // ======= StartPreviewLoop（表示用ループ、ログ追加済み） =======
        // StartPreviewLoop - 表示用ループ（ログ追加済み）
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
                            Bitmap? toShow = null;
                            try
                            {
                                if (_selectedRegion.Width > 0 && _selectedRegion.Height > 0)
                                {
                                    var safeR = new System.Drawing.Rectangle(
                                        Math.Max(0, Math.Min(_selectedRegion.X, bmp.Width - 1)),
                                        Math.Max(0, Math.Min(_selectedRegion.Y, bmp.Height - 1)),
                                        Math.Max(1, Math.Min(_selectedRegion.Width, Math.Max(1, bmp.Width - _selectedRegion.X))),
                                        Math.Max(1, Math.Min(_selectedRegion.Height, Math.Max(1, bmp.Height - _selectedRegion.Y)))
                                    );

                                    var crop = new Bitmap(safeR.Width, safeR.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                                    using (var g = Graphics.FromImage(crop))
                                    {
                                        g.DrawImage(bmp, new System.Drawing.Rectangle(0, 0, crop.Width, crop.Height), safeR, GraphicsUnit.Pixel);
                                    }
                                    toShow = crop;
                                }
                                else
                                {
                                    var full = new Bitmap(bmp.Width, bmp.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                                    using (var g = Graphics.FromImage(full))
                                    {
                                        g.DrawImage(bmp, 0, 0, bmp.Width, bmp.Height);
                                    }
                                    toShow = full;
                                }

                                string hash = ComputeSimpleHash(toShow);
                                Debug.WriteLine($"[PreviewHash] computed={hash} prev={_previewLastHash} region={_selectedRegion.X},{_selectedRegion.Y},{_selectedRegion.Width}x{_selectedRegion.Height}");
                                Debug.WriteLine($"[PreviewImageInfo] toShowSize={toShow.Width}x{toShow.Height} checksum={ComputeQuickChecksum(toShow)} hash={hash}");

                                if (_previewLastHash != null && _previewLastHash == hash)
                                {
                                    Debug.WriteLine("[Preview] hash unchanged -> skip UI update");
                                    try { await Task.Delay(_previewIntervalMs, ct).ConfigureAwait(false); } catch (TaskCanceledException) { break; }
                                    continue;
                                }

                                _previewLastHash = hash;

                                BitmapSource? bs = null;
                                try
                                {
                                    bs = BitmapToBitmapSource(toShow);
                                    if (bs != null && bs.CanFreeze) bs.Freeze();
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"[PreviewLoop] BitmapToBitmapSource failed: {ex}");
                                    try { await Task.Delay(_previewIntervalMs, ct).ConfigureAwait(false); } catch (TaskCanceledException) { break; }
                                    continue;
                                }

                                UpdatePreview(bs);
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[PreviewLoop] inner processing exception: {ex}");
                            }
                            finally
                            {
                                try { toShow?.Dispose(); } catch (Exception ex) { Debug.WriteLine($"[PreviewLoop] dispose error: {ex}"); }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[PreviewLoop] Exception: {ex}");
                    }

                    try
                    {
                        await Task.Delay(_previewIntervalMs, ct).ConfigureAwait(false);
                    }
                    catch (TaskCanceledException)
                    {
                        break;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[PreviewLoop] Delay exception: {ex}");
                    }
                }

                Debug.WriteLine("[PreviewLoop] stopped");
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
                    txtNotificationContent?.Focus();
                    return;
                }

                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                bool ok = false;

                try
                {
                    ok = await NotifyDiscordAsync(webhookUrl, userText, cts.Token);
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
                    try
                    {
                        txtNotificationContent.Clear();
                        txtNotificationContent.Focus();
                    }
                    catch { }
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

        private string GetWebhookUrl(string key = "ClipDiscordApp")
        {
            try
            {
                if (string.IsNullOrWhiteSpace(key)) key = "ClipDiscordApp";

                var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                var appDataDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ClipDiscordApp");

                var candidates = new[]
                {
                    Path.Combine(exeDir, "discord_webhooks.json"),
                    Path.Combine(appDataDir, "discord_webhooks.json")
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

                // 環境変数フォールバック（大文字キー）
                var envKey = $"DISCORD_WEBHOOK_{key.ToUpperInvariant()}";
                var envUrl = Environment.GetEnvironmentVariable(envKey);
                if (!string.IsNullOrWhiteSpace(envUrl) && Uri.TryCreate(envUrl.Trim(), UriKind.Absolute, out var envUri) && (envUri.Scheme == "https" || envUri.Scheme == "http"))
                {
                    Debug.WriteLine($"[GetWebhookUrl] loaded webhook from ENV {envKey}");
                    return envUrl.Trim();
                }

                // 追加フォールバック
                var fallbackEnv = Environment.GetEnvironmentVariable("DISCORD_WEBHOOK_CLIPDISCORDAPP");
                if (!string.IsNullOrWhiteSpace(fallbackEnv) && Uri.TryCreate(fallbackEnv.Trim(), UriKind.Absolute, out var fbUri) && (fbUri.Scheme == "https" || fbUri.Scheme == "http"))
                {
                    Debug.WriteLine("[GetWebhookUrl] loaded webhook from ENV DISCORD_WEBHOOK_CLIPDISCORDAPP");
                    return fallbackEnv.Trim();
                }

                Debug.WriteLine("[GetWebhookUrl] webhook not found in any candidate locations");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetWebhookUrl] unexpected error: {ex}");
            }

            return string.Empty;
        }

        private void BtnStartMonitoring_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1) ハッシュとフラグの初期化は任意（元の挙動に合わせるならここは空にしても良い）
                // _previewLastHash = null;
                // _lastPreviewHash = null;
                // _skipFirstFrameAfterStart = true;
                // _waitingForNextFrameAfterNotify = false;

                // 2) 領域選択を先に行う（プレビューは領域確定後に開始）
                var selector = new RegionSelectorWindow { Owner = this };
                var dlg = selector.ShowDialog();
                if (dlg != true)
                {
                    StatusText.Text = "Monitoring cancelled (no region selected)";
                    return;
                }

                // 3) 選択領域を内部座標に変換して保存・表示
                var d = selector.SelectedRegion;
                var selRect = new System.Windows.Rect(d.X, d.Y, d.Width, d.Height);
                _selectedRegion = ConvertImageRectToBitmapRect(selRect, CapturedImage);
                DrawSelectionOnOverlay(_selectedRegion);

                // 4) プレビューループ開始（領域が確定してから）
                // 元の挙動に戻すなら StartPreviewLoop を呼ぶか呼ばないかは既存実装に合わせる。
                StartPreviewLoop();

                // 5) 即時スナップショットで UI を更新（Freeze は残すと安全）
                try
                {
                    using var full = CaptureFrameBitmap();
                    using var crop = CropToRegion(full, _selectedRegion);
                    var bs = BitmapToBitmapSource(crop);
                    if (bs.CanFreeze) bs.Freeze();

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            CapturedImage.Source = bs;
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

                // 6) UI ボタン状態とステータス更新
                BtnStartMonitoring.IsEnabled = false;
                BtnStopMonitoring.IsEnabled = true;
                StatusText.Text = "Monitoring...";

                // 7) 既存の監視をキャンセルして新規作成（例外はログに残す）
                try
                {
                    _monitoringCts?.Cancel();
                    _monitoringCts?.Dispose();
                    _monitoringCts = null;
                    Debug.WriteLine("[MonitorCancel] monitoring cancelled (pre-start)");
                }
                catch (ObjectDisposedException ex)
                {
                    Debug.WriteLine($"[MonitorCancel] ObjectDisposedException while cancelling monitoring: {ex}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[MonitorCancel] unexpected error while cancelling monitoring: {ex}");
                }
                _monitoringCts = new CancellationTokenSource();

                // 8) MonitorService の初期化（元のコールバックと互換ハッシュを使う）
                if (_monitorService == null)
                {
                    _monitorService = new MonitorService(
                        _logDir,
                        _tesseractEngine,
                        () => CaptureFrameBitmap(),
                        (bmp, rect) => CropToRegion(bmp, rect),
                        // 元の呼び出しを壊さないために ComputeSimpleHash を渡す（互換ラッパーがあればそれでも可）
                        (bmp) => ComputeSimpleHash(bmp),
                        (bmp) => BitmapToBitmapSource(bmp),
                        async (object m) => { await HandleMatchAsync((ExtractMatch)m); },
                        () => LoadRules(),
                        // 元の単純なプレビューコールバックに戻す（フラグチェックは削除）
                        (bmpSource) =>
                        {
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

                // 9) 監視ループをバックグラウンドで開始
                _ = Task.Run(async () =>
                {
                    try
                    {
                        await StartMonitoringAsync(_selectedRegion, _monitoringCts.Token);
                    }
                    catch (OperationCanceledException)
                    {
                        // 正常なキャンセル
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
                            if (_monitoringCts == null || !_monitoringCts.IsCancellationRequested) StatusText.Text = "Stopped";
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[BtnStartMonitoring_Click] unexpected: {ex}");
                System.Windows.MessageBox.Show($"開始に失敗しました: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private async Task<bool> NotifyDiscordAsync(string webhookUrl, string content, CancellationToken ct, int maxRetries = 3)
        {
            if (string.IsNullOrWhiteSpace(webhookUrl))
            {
                Debug.WriteLine("[Discord] webhook empty");
                return false;
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                Debug.WriteLine("[Discord] content empty -> Discord will ignore");
                return false;
            }

            var payload = new { content = content.Trim() };
            var json = JsonSerializer.Serialize(payload);
            var attempt = 0;
            var backoff = TimeSpan.FromSeconds(1);
            var maxBackoff = TimeSpan.FromSeconds(30);

            while (attempt < maxRetries)
            {
                attempt++;
                using var body = new StringContent(json, Encoding.UTF8, "application/json");

                try
                {
                    using var resp = await _discordHttpClient.PostAsync(webhookUrl, body, ct);
                    var status = (int)resp.StatusCode;
                    var respText = await resp.Content.ReadAsStringAsync(ct);
                    Debug.WriteLine($"[Discord] attempt={attempt} status={status} body={respText}");

                    if (resp.IsSuccessStatusCode)
                        return true;

                    if (status == 401)
                    {
                        Debug.WriteLine("[Discord] Unauthorized 401 - webhook invalid");
                        return false;
                    }

                    if (status == 429)
                    {
                        // …既存の retry_after 処理…
                    }

                    if (status >= 400 && status < 500)
                    {
                        Debug.WriteLine($"[Discord] client error {status}: {respText}");
                        return false;
                    }

                    Debug.WriteLine($"[Discord] server error {status}: {respText}");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[Discord] send exception: {ex}");
                }

                await Task.Delay(backoff, ct);
                backoff = backoff < maxBackoff ? TimeSpan.FromSeconds(Math.Min(backoff.TotalSeconds * 2, maxBackoff.TotalSeconds)) : maxBackoff;
            }

            Debug.WriteLine("[Discord] send failed after retries");
            return false;
        }

        private DateTime GetTokyoNow()
        {
            try
            {
                TimeZoneInfo tz = null;

                try
                {
                    tz = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[GetTokyoNow] Windows TZ not found: {ex.Message}");
                }

                if (tz == null)
                {
                    try
                    {
                        tz = TimeZoneInfo.FindSystemTimeZoneById("Asia/Tokyo");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[GetTokyoNow] IANA TZ not found: {ex.Message}");
                    }
                }

                if (tz != null)
                {
                    return TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, tz);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetTokyoNow] unexpected: {ex}");
            }

            return DateTime.Now;
        }

        // Logs: 古いファイル削除
        private void CleanupOldDebugFiles(string dirPath, int keepDays)
        {
            try
            {
                if (!Directory.Exists(dirPath)) return;
                var cutoff = DateTime.Now.Date.AddDays(-keepDays);
                foreach (var file in Directory.EnumerateFiles(dirPath, "*.png"))
                {
                    try
                    {
                        var fi = new FileInfo(file);
                        if (fi.LastWriteTime.Date < cutoff)
                        {
                            fi.Delete();
                            Debug.WriteLine($"[Cleanup] Deleted {fi.Name}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[Cleanup] Failed to delete {file}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Cleanup] CleanupOldDebugFiles failed: {ex.Message}");
            }
        }

        // 日次クリーンアップ
        private void StartDailyCleanupTask()
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    while (true)
                    {
                        var now = DateTime.Now;
                        var nextMidnight = now.Date.AddDays(1);
                        var delay = nextMidnight - now;
                        await Task.Delay(delay);

                        CleanupOldDebugFiles(_logDir, _debugKeepDays);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[CleanupTask] Exception: {ex}");
                }
            });
        }

        // XAML Event Handlers (UI binding)
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
                // 注意: CapturedImage が拡大縮小されている場合は、変換が必要
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

        private void UpdatePreview(BitmapSource bs)
        {
            if (bs == null) return;

            // Freeze は呼び出し元でも行っている想定だが冗長にチェック
            if (bs.CanFreeze) bs.Freeze();

            lock (_previewLock)
            {
                // スロットリング
                if ((DateTime.UtcNow - _lastPreviewUpdate) < _minPreviewInterval) return;
                _lastPreviewUpdate = DateTime.UtcNow;

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        CapturedImage.Source = bs;
                        var src = PresentationSource.FromVisual(this);
                        double dpiScaleX = src?.CompositionTarget?.TransformToDevice.M11 ?? 1.0;
                        double dpiScaleY = src?.CompositionTarget?.TransformToDevice.M22 ?? 1.0;
                        OverlayCanvas.Width = Math.Round(bs.PixelWidth / dpiScaleX);
                        OverlayCanvas.Height = Math.Round(bs.PixelHeight / dpiScaleY);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[UpdatePreview] UI set failed: {ex}");
                    }
                }), DispatcherPriority.Render);
            }
        }
        private MatchResultItem ToMatchResultItem(object extractMatch)
        {
            if (extractMatch == null) return null;
            // 可能なら実際の型名を使ってキャストしてください（例: ClipDiscordApp.Models.ExtractMatch）
            try
            {
                dynamic em = extractMatch;
                string id = em.RuleId ?? em.Id ?? string.Empty;
                string name = em.RuleName ?? em.Name ?? string.Empty;
                string[] matches = null;
                try { matches = em.Matches as string[] ?? new string[0]; } catch { matches = new string[0]; }
                return new MatchResultItem(id, name, matches);
            }
            catch
            {
                return null;
            }
        }
    }
}