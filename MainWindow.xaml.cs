using ClipDiscordApp.Models;
using ClipDiscordApp.Parsers;
using ClipDiscordApp.Services;
using ClipDiscordApp.Utils;
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
    public partial class MainWindow : Window
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
        // フィールド（MainWindow クラス内に追加してください）
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
                // 例: _tesseractEngine = new TesseractEngine(tessDataDir, "eng+jpn", EngineMode.Default)
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
        // StartMonitoringAsync の全文（差分検出と前処理向上を含む）
        public async Task StartMonitoringAsync(System.Drawing.Rectangle initialRegion, CancellationToken cancellationToken)
        {
            var rules = LoadRules();

            // 新: ClipDiscordApp.Utils のプリセットを使う
            var prepParams = OcrPreprocessorPresets.Conservative;

            Directory.CreateDirectory(_logDir);

            // 閾値（必要に応じて調整）
            const double TEMPLATE_SCORE_THRESHOLD = 0.80;
            const double NOTIFY_THRESHOLD = 0.75;
            const double LOG_ONLY_THRESHOLD = 0.55;

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

                    // デバッグ保存（任意）
                    try
                    {
                        var fname = Path.Combine(_logDir, $"pre_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}.png");
                        prepped.Save(fname);
                        Debug.WriteLine($"[OCR] Preprocessed image saved: {fname}");
                    }
                    catch (Exception ex) { Debug.WriteLine($"[OCR] Save prepped failed: {ex.Message}"); }

                    // ----- 1) テンプレートマッチング（優先） -----
                    // TemplateMatcher.Check は (found:bool, label:string, score:double) を返す想定
                    var templateResult = TemplateMatcher.Check(prepped); // implement elsewhere
                    if (templateResult.found)
                    {
                        Debug.WriteLine($"[Template] found {templateResult.label} score={templateResult.score}");
                        if (templateResult.score >= TEMPLATE_SCORE_THRESHOLD)
                        {
                            var toParse = templateResult.label; // "BUY" or "SELL"
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
                            continue; // 次ループへ
                        }
                        else
                        {
                            Debug.WriteLine($"[Template] low score {templateResult.score}, fall back to OCR");
                        }
                    }

                    // ----- 2) フォールバック: OCR(binary と gray の両方) -----
                    string textBin = string.Empty;
                    string textGray = string.Empty;

                    try
                    {
                        var up = (textBin ?? string.Empty).ToUpperInvariant();
                        var m = System.Text.RegularExpressions.Regex.Match(up, @"\b(BUY|SELL)\b");
                        if (m.Success)
                        {
                            var label = m.Groups[1].Value; // "BUY" or "SELL"
                            Debug.WriteLine($"[OcrHelpers] Regex immediate match: {label} from '{up}'");

                            // ルールパースしてハンドラへ渡す（TemplateMatcher と同等の早期処理）
                            var matches = OcrParser.ParseByRules(label, rules);
                            if (matches != null && matches.Any())
                            {
                                foreach (var mm in matches)
                                {
                                    var key = !string.IsNullOrEmpty(mm.RuleId) ? mm.RuleId : mm.RuleName;
                                    if (string.IsNullOrEmpty(key)) key = Guid.NewGuid().ToString();

                                    if (_lastNotified.TryGetValue(key, out DateTime last) && (DateTime.UtcNow - last) < _notifyCooldown)
                                    {
                                        Debug.WriteLine($"[OCR] Skip notify for {key} (cooldown)");
                                        continue;
                                    }

                                    _lastNotified[key] = DateTime.UtcNow;
                                    _ = HandleMatchAsync(mm); // 非同期で実行（既存ハンドラが Task を返す想定）
                                    Debug.WriteLine($"[OCR] Match(from regex): rule={mm.RuleName} matches=[{string.Join(',', mm.Matches)}]");
                                }
                            }

                            // 正規表現で判定できたので以降の OCR 正規化／スコア処理はスキップして次ループへ
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[OcrHelpers] Regex immediate match processing failed: {ex}");
                    }


                    try
                    {
                        // Gray bitmap作成用ヘルパ。実装がない場合は prepped を直接使うか
                        using var grayBmp = OcrPreprocessor.GetGrayBitmap(prepped); // implement if needed
                        using var pixGray = PixConverter.ToPix(grayBmp);
                        textGray = OcrHelpers.DoOcrWithRetries(_tesseractEngine, pixGray) ?? string.Empty;
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[OCR] gray OCR failed: {ex.Message}");
                    }

                    Debug.WriteLine($"[OCR] Raw bin len={textBin.Length} bin:'{(textBin.Length > 200 ? textBin.Substring(0, 200) + "..." : textBin)}'");
                    Debug.WriteLine($"[OCR] Raw gray len={textGray.Length} gray:'{(textGray.Length > 200 ? textGray.Substring(0, 200) + "..." : textGray)}'");

                    // ----- 3) 候補から BUY/SELL を両方評価して最良ラベルを選ぶ -----
                    // NormalizeAndDetectLabels は rawCandidates から複数 LabelCandidate を返す想定
                    // LabelCandidate: { string Text; string Label; double Confidence; string Source; }
                    var rawCandidates = new List<(string text, string source)>();
                    rawCandidates.Add((textBin, "binary"));
                    rawCandidates.Add((textGray, "gray"));

                    var labelCandidates = new List<LabelCandidate>();
                    foreach (var (txt, src) in rawCandidates)
                    {
                        if (string.IsNullOrWhiteSpace(txt)) continue;
                        try
                        {
                            var list = OcrHelpers.NormalizeAndDetectLabels(new[] { txt }, src); // implement to return List<LabelCandidate>
                            if (list != null) labelCandidates.AddRange(list);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[OcrHelpers] NormalizeAndDetectLabels failed for {src}: {ex.Message}");
                        }
                    }

                    // best per label
                    var bestSell = labelCandidates.Where(c => c.Label == "SELL").OrderByDescending(c => c.Confidence).FirstOrDefault();
                    var bestBuy = labelCandidates.Where(c => c.Label == "BUY").OrderByDescending(c => c.Confidence).FirstOrDefault();

                    var chosen = (bestSell?.Confidence ?? 0) >= (bestBuy?.Confidence ?? 0) ? bestSell : bestBuy;

                    Debug.WriteLine($"[OCR] bestSell={bestSell?.Text}:{bestSell?.Confidence} bestBuy={bestBuy?.Text}:{bestBuy?.Confidence} chosen={chosen?.Label}:{chosen?.Text}:{chosen?.Confidence}");

                    // ログのみ or 通知判定
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

                                if (_lastNotified.TryGetValue(key, out var last) && (DateTime.UtcNow - last) < _notifyCooldown)
                                {
                                    Debug.WriteLine($"[OCR] Skip notify for {key} (cooldown)");
                                    continue;
                                }

                                _lastNotified[key] = DateTime.UtcNow;
                                _ = Task.Run(() => HandleMatchAsync(m));
                                Debug.WriteLine($"[OCR] Match: rule={m.RuleName} matches=[{string.Join(',', m.Matches)}]");
                            }
                        }
                    }
                    else if (chosen != null && chosen.Confidence >= LOG_ONLY_THRESHOLD)
                    {
                        Debug.WriteLine($"[OCR] Low-confidence candidate (log only): {chosen.Label} '{chosen.Text}' conf={chosen.Confidence}");
                    }
                    else
                    {
                        Debug.WriteLine("[OCR] No confident candidate found");
                    }

                    // ----- 4) プレビュー更新（既存ロジック） -----
                    try
                    {
                        Bitmap bmpClone = null;
                        try
                        {
                            bmpClone = (Bitmap)roi.Clone();
                            var hash = ComputeSimpleHash(bmpClone);
                            if (hash != _lastPreviewHash)
                            {
                                _lastPreviewHash = hash;

                                BitmapSource bmpSource = null;
                                try
                                {
                                    bmpSource = BitmapToBitmapSource(bmpClone);
                                    if (bmpSource != null && bmpSource.CanFreeze) bmpSource.Freeze();
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"[Preview] BitmapToBitmapSource failed: {ex.Message}");
                                    bmpSource = null;
                                }

                                if (bmpSource != null)
                                {
                                    Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        try
                                        {
                                            CapturedImage.Source = bmpSource;
                                        }
                                        catch (Exception ex)
                                        {
                                            Debug.WriteLine($"[Preview] Update failed in Dispatcher: {ex.Message}");
                                        }
                                    }));
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
                        Debug.WriteLine($"[Preview] Error computing/updating preview: {ex}");
                    }

                    await Task.Delay(200, cancellationToken);
                }
                catch (TaskCanceledException tex)
                {
                    Debug.WriteLine($"[StartMonitoring] TaskCanceledException: {tex.Message}\n{tex.StackTrace}");
                    break;
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[StartMonitoring] Error: {ex}");
                    await Task.Delay(500, cancellationToken);
                }
            }
        }

        // シンプルな差分ハッシュ（高速で概ねの変化検出用）
        private string ComputeSimpleHash(Bitmap bmp)
        {
            // サンプリング間隔は画像サイズに応じて調整（高速化のため一部ピクセルを使う）
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
            // ログディレクトリを決める
            _logsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            Directory.CreateDirectory(_logsDir);

            // NOTE: _tesseractEngine はコンストラクタで初期化済み（再代入しない）
            // TemplateMatcher 初期化: Templates フォルダに BUY/SELL を配置しておく
            var templatesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            TemplateMatcher.Initialize(templatesPath);

            // MonitorService の生成と Start は BtnStartMonitoring_Click 側で行うためここでは行わない
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

                // _selectedRegion が初期化済みであればログする（多くは未選択なので幅高さ=0）
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
                var contentText = $"{now:yyyy-MM-dd HH:mm:ss} SELL" +
                    $"";

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
            // 常に選択ダイアログを表示して選択済み領域を取得する
            var selector = new RegionSelectorWindow { Owner = this };
            var dlg = selector.ShowDialog();
            if (dlg != true)
            {
                StatusText.Text = "Monitoring cancelled (no region selected)";
                return;
            }
            // selector.SelectedRegion は ClipDiscordApp.Models.DrawingRect 型
            var d = selector.SelectedRegion;
            var selRect = new System.Windows.Rect(d.X, d.Y, d.Width, d.Height);
            _selectedRegion = ConvertImageRectToBitmapRect(selRect, CapturedImage);
            DrawSelectionOnOverlay(_selectedRegion);

            // UI 更新
            BtnStartMonitoring.IsEnabled = false;
            BtnStopMonitoring.IsEnabled = true;
            StatusText.Text = "Monitoring...";

            // 既に実行中なら古い CTS をキャンセルしておく
            try { _monitoringCts?.Cancel(); } catch { }

            // 新規に CTS を作成（既存フィールド名に合わせる）
            _monitoringCts = new CancellationTokenSource();

            // MonitorService が未生成ならここで生成する（毎回新しく作ってもOK）
            if (_monitorService == null)
            {
                _monitorService = new MonitorService(
                    _logsDir,
                    _tesseractEngine,
                    // captureFrame
                    () => CaptureFrameBitmap(),
                    // cropToRegion
                    (bmp, rect) => CropToRegion(bmp, rect),
                    // computeHash
                    (bmp) => ComputeSimpleHash(bmp),
                    // bmpToSource
                    (bmp) => BitmapToBitmapSource(bmp),
                    // handleMatchAsync: MonitorService expects Func<object, Task> (object-based).
                    // Cast object -> ExtractMatch inside this lambda before calling your existing handler.
                    async (object m) => { await HandleMatchAsync((ExtractMatch)m); },
                    // loadRules
                    () => LoadRules(),
                    // setPreviewAction: UI スレッド経由で CapturedImage.Source にセット
                    (bmpSource) =>
                    {
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            try { CapturedImage.Source = bmpSource; } catch (Exception ex) { Debug.WriteLine($"[MainWindow] preview set failed: {ex.Message}"); }
                        }));
                    });
            }

            // StartMonitoring をバックグラウンドで開始
            _ = Task.Run(async () =>
            {
                try
                {
                    await StartMonitoringAsync(_selectedRegion, _monitoringCts.Token);
                }
                catch (OperationCanceledException)
                {
                    // 正常停止
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[BtnStart] Error: {ex}");
                    Dispatcher.Invoke(() => StatusText.Text = "Error");
                }
                finally
                {
                    // UI を元に戻す
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
        // selRectInImage: RegionSelectorWindow.SelectedRegion が System.Windows.Rect を返す想定
        // imageControl: プレビュー表示に使っている Image コントロール（例: CapturedImage）
        private System.Drawing.Rectangle ConvertImageRectToBitmapRect(System.Windows.Rect selRectInImage, System.Windows.Controls.Image imageControl)
        {
            if (imageControl?.Source is System.Windows.Media.Imaging.BitmapSource bmp)
            {
                // Image コントロールの表示サイズ（DIP: WPF単位）
                double displayW = imageControl.ActualWidth;
                double displayH = imageControl.ActualHeight;

                // Bitmap のピクセルサイズ
                double sourceW = bmp.PixelWidth;
                double sourceH = bmp.PixelHeight;

                // DPI / スケール補正
                var src = PresentationSource.FromVisual(this);
                double dpiScaleX = 1.0, dpiScaleY = 1.0;
                if (src?.CompositionTarget != null)
                {
                    dpiScaleX = src.CompositionTarget.TransformToDevice.M11;
                    dpiScaleY = src.CompositionTarget.TransformToDevice.M22;
                }

                // 表示ピクセルに換算（DIP -> physical pixels）
                double displayPxW = displayW * dpiScaleX;
                double displayPxH = displayH * dpiScaleY;

                // 元画像ピクセルにマッピングするスケール（displayPx -> source pixels）
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

            // フォールバック（SelectedRegion がすでにスクリーン/ピクセル座標の場合）
            return new System.Drawing.Rectangle(
                (int)Math.Max(0, selRectInImage.X),
                (int)Math.Max(0, selRectInImage.Y),
                (int)Math.Max(1, selRectInImage.Width),
                (int)Math.Max(1, selRectInImage.Height)
            );
        }
    }
}