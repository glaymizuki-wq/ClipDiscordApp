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

        // Selection UI helpers
        private System.Windows.Point? _dragStart;
        private System.Windows.Shapes.Rectangle? _selectionRectVisual;

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
                        using var frame = CaptureFrameBitmap();
                        if (frame == null)
                        {
                            await Task.Delay(200, cancellationToken);
                            continue;
                        }

                        using var roi = CropToRegion(frame, initialRegion);
                        using var prepped = OcrPreprocessor.Preprocess(roi, prepParams);

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
                } // end while
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

        private async Task HandleMatchAsync(ExtractMatch match)
        {
            try
            {
                Debug.WriteLine($"[HandleMatchAsync] Handling match rule={match.RuleName} text={string.Join(",", match.Matches)}");
                await Task.Delay(1);
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
            _monitoringCts?.Cancel();
            BtnStartMonitoring.IsEnabled = true;
            BtnStopMonitoring.IsEnabled = false;
            StatusText.Text = "Stopped";
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
            try
            {
                var webhookUrl = GetWebhookUrl("ClipDiscordApp");
                if (string.IsNullOrWhiteSpace(webhookUrl))
                {
                    System.Windows.MessageBox.Show("Webhook URL が設定されていません。", "設定エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var now = DateTime.Now;
                var contentText = $"{now:yyyy-MM-dd HH:mm:ss} SELL";
                var payload = new { content = contentText };
                var json = JsonSerializer.Serialize(payload);
                using var httpContent = new StringContent(json, Encoding.UTF8, "application/json");

                var resp = await _httpClient.PostAsync(webhookUrl, httpContent);
                if (!resp.IsSuccessStatusCode)
                {
                    Debug.WriteLine($"[BtnTestNotification] Discord送信失敗: {(int)resp.StatusCode} {resp.ReasonPhrase}");
                    System.Windows.MessageBox.Show($"送信失敗: {(int)resp.StatusCode} {resp.ReasonPhrase}", "送信エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    Debug.WriteLine("[BtnTestNotification] Discord送信成功");
                    System.Windows.MessageBox.Show("送信完了", "通知", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[BtnTestNotification] 例外: {ex}");
                System.Windows.MessageBox.Show($"例外が発生しました: {ex.Message}", "例外", MessageBoxButton.OK, MessageBoxImage.Error);
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
                var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "discord_webhooks.json");
                if (!File.Exists(configPath))
                {
                    Debug.WriteLine($"[GetWebhookUrl] 設定ファイルが見つかりません: {configPath}");
                    return string.Empty;
                }

                var json = File.ReadAllText(configPath, Encoding.UTF8);
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty(key, out var val))
                {
                    var url = val.GetString();
                    if (!string.IsNullOrWhiteSpace(url)) return url.Trim();
                }
                else
                {
                    Debug.WriteLine($"[GetWebhookUrl] キー '{key}' が設定ファイルに見つかりません");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetWebhookUrl] 読み込みエラー: {ex.Message}");
            }

            return string.Empty;
        }

        private void BtnStartMonitoring_Click(object sender, RoutedEventArgs e)
        {
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
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            try { CapturedImage.Source = bmpSource; } catch (Exception ex) { Debug.WriteLine($"[MainWindow] preview set failed: {ex.Message}"); }
                        }));
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
    }
}