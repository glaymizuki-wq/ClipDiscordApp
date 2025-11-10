using ClipDiscordApp.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms; // Screen
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

        private CancellationTokenSource? _monitoringCts;
        private System.Drawing.Rectangle _selectedRegion = new System.Drawing.Rectangle(100, 100, 400, 100);
        private readonly string _rulesFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rules.json");
        private readonly string _logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ocr_logs");
     
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
            var tessDataDir = tessDir;
            try
            {
                _tesseractEngine = new TesseractEngine(tessDataDir, "jpn", EngineMode.Default);
                _tesseractEngine.SetVariable("tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789:");
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

            var prepParams = new OcrPreprocessor.Params
            {
                Scale = 2,
                AdaptiveBlockSize = 31,
                AdaptiveC = 8,
                UseClahe = true,
                Sharpen = false,
                MorphKernel = 2
            };

            Directory.CreateDirectory(_logDir);

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

                    // OCR 実行
                    string ocrText;
                    try
                    {
                        using var pix = PixConverter.ToPix(prepped);
                        using var page = _tesseractEngine.Process(pix);
                        ocrText = page.GetText()?.Trim() ?? string.Empty;
                    }
                    catch
                    {
                        using var pix = OcrPreprocessor.PreprocessToPix(prepped);
                        using var page = _tesseractEngine.Process(pix);
                        ocrText = page.GetText()?.Trim() ?? string.Empty;
                    }

                    Debug.WriteLine($"[OCR] Raw length={ocrText.Length} Raw:'{(ocrText.Length > 200 ? ocrText.Substring(0, 200) + "..." : ocrText)}'");

                    var matches = OcrParser.ParseByRules(ocrText, rules);

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

                    // プレビュー差し替えは「変化があった場合のみ」
                    try
                    {
                        using var bmpClone = (Bitmap)roi.Clone();
                        var hash = ComputeSimpleHash(bmpClone);
                        if (hash != _lastPreviewHash)
                        {
                            _lastPreviewHash = hash;
                            Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    CapturedImage.Source = BitmapToBitmapSource(bmpClone);
                                }
                                catch (Exception ex)
                                {
                                    Debug.WriteLine($"[Preview] Update failed: {ex.Message}");
                                }
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[Preview] Error computing/updating preview: {ex}");
                    }

                    await Task.Delay(200, cancellationToken);
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
            StatusText.Text = "Ready";
            // Draw initial selection if needed
            DrawSelectionOnOverlay(_selectedRegion);
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
            _selectedRegion = ConvertDrawingRectToRectangle(selector.SelectedRegion);
            DrawSelectionOnOverlay(_selectedRegion);

            BtnStartMonitoring.IsEnabled = false;
            BtnStopMonitoring.IsEnabled = true;
            StatusText.Text = "Monitoring...";

            _monitoringCts = new CancellationTokenSource();
            _ = Task.Run(async () =>
            {
                try
                {
                    await StartMonitoringAsync(_selectedRegion, _monitoringCts.Token);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[BtnStart] Error: {ex}");
                    Dispatcher.Invoke(() => StatusText.Text = "Error");
                }
            });
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

        private void BtnTestNotification_Click(object sender, RoutedEventArgs e)
        {
            var testMatch = new ExtractMatch
            {
                RuleId = "test",
                RuleName = "通知テスト",
                Matches = new List<string> { "TEST123" }
            };
            _ = HandleMatchAsync(testMatch);
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
    }
}