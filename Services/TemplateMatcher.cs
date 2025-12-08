// TemplateMatcher.cs
// Requires: OpenCvSharp, OpenCvSharp.Extensions
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System.Drawing;

public class MatchResult
{
    public bool Found { get; set; }
    public string Label { get; set; }
    public double BestScore { get; set; }
    public int TriedCandidates { get; set; }
    public long ElapsedMs { get; set; }
}

public static class TemplateMatcher
{
    // テンプレートを保持するエントリ（Mat と元ファイル名を保持）
    private class TemplateEntry
    {
        public Mat Mat { get; set; }
        public string FileName { get; set; }
    }

    // label -> list of template entries (grayscale, preprocessed)
    private static readonly Dictionary<string, List<TemplateEntry>> _templates = new();

    // Configurable options
    public static double AcceptThreshold { get; set; } = 0.80; // default threshold (tuneable)
    public static double[] ScaleCandidates { get; set; } = new[] { 0.98, 1.0, 1.02 }; // reduce candidates for speed
    public static int MaxAllowedMsBeforeCheck { get; set; } = 500; // check cancellation every N ms

    // Optional debug output directory (used only in DEBUG)
    private static readonly string _debugOutDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

    /// <summary>
    /// Initialize: load templates from folder (file names should include label as second segment e.g. tpl_buy_01.png)
    /// </summary>
    public static void Initialize(string templatesDir)
    {
        // Dispose existing Mats if any
        foreach (var kv in _templates)
        {
            foreach (var e in kv.Value)
            {
                try { e.Mat?.Dispose(); } catch { }
            }
        }
        _templates.Clear();

        if (!Directory.Exists(templatesDir))
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Templates dir not found: {templatesDir}");
            return;
        }

        var loadedFiles = new List<string>();
        foreach (var f in Directory.EnumerateFiles(templatesDir, "*.png"))
        {
            var name = Path.GetFileNameWithoutExtension(f).ToLowerInvariant();
            var parts = name.Split('_'); // expected tpl_buy_01
            if (parts.Length < 2) continue;
            var label = parts[1]; // buy, sell, etc.
            try
            {
                using var b = new System.Drawing.Bitmap(f);
                var mat = BitmapConverter.ToMat(b);
                Mat gray = new Mat();
                if (mat.Channels() == 3 || mat.Channels() == 4)
                    Cv2.CvtColor(mat, gray, ColorConversionCodes.BGR2GRAY);
                else
                    gray = mat.Clone();

                // optional normalization: resize very large templates down to reasonable size
                if (gray.Width > 1000 || gray.Height > 300)
                {
                    var scaled = new Mat();
                    double scale = Math.Min(1000.0 / gray.Width, 300.0 / gray.Height);
                    Cv2.Resize(gray, scaled, new OpenCvSharp.Size(0, 0), scale, scale, InterpolationFlags.Area);
                    gray.Dispose();
                    gray = scaled;
                }

                if (!_templates.TryGetValue(label, out var list))
                {
                    list = new List<TemplateEntry>();
                    _templates[label] = list;
                }

                // 保存は TemplateEntry として行う（Mat の所有権はここで保持）
                list.Add(new TemplateEntry { Mat = gray, FileName = Path.GetFileName(f) });
                loadedFiles.Add(Path.GetFileName(f));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Load failed {f}: {ex.Message}");
            }
        }

        System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Loaded labels: {_templates.Keys.Count}");
#if DEBUG
        System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] Template labels: {string.Join(", ", _templates.Keys)}");
#endif
        if (loadedFiles.Count > 0)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Found files: {string.Join(", ", loadedFiles)}");
        }
    }

    /// <summary>
    /// Normalize template-derived label into canonical rule label (SELL/BUY) when possible.
    /// </summary>
    private static string NormalizeTemplateLabelToRule(string label)
    {
        if (string.IsNullOrEmpty(label)) return null;
        var n = label.Trim().ToUpperInvariant();
        if (n.Contains("DOWN") || n.Contains("SELL")) return "SELL";
        if (n.Contains("UP") || n.Contains("BUY")) return "BUY";
        return n; // fallback: return original normalized label
    }

    /// <summary>
    /// Async Check that respects cancellation and logs summary
    /// preppedBmp must be preprocessed with the same pipeline used to create templates (grayscale/resized)
    /// </summary>
    public static async Task<MatchResult> CheckAsync(System.Drawing.Bitmap preppedBmp, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var swTotal = System.Diagnostics.Stopwatch.StartNew();
            int triedCandidates = 0;
            double bestScore = double.MinValue;
            string bestLabel = null;
            string bestFile = null;

            if (_templates.Count == 0)
            {
                swTotal.Stop();
                System.Diagnostics.Debug.WriteLine("[TemplateMatcher] No templates loaded, skipping matching");
                return new MatchResult
                {
                    Found = false,
                    Label = null,
                    BestScore = 0.0,
                    TriedCandidates = 0,
                    ElapsedMs = swTotal.ElapsedMilliseconds
                };
            }

            using var roiMatColor = BitmapConverter.ToMat(preppedBmp);
            Mat roi = new Mat();
            if (roiMatColor.Channels() == 3 || roiMatColor.Channels() == 4)
                Cv2.CvtColor(roiMatColor, roi, ColorConversionCodes.BGR2GRAY);
            else
                roi = roiMatColor.Clone();

            try
            {
                // Quick filter to skip heavy matching when ROI clearly doesn't match template color/brightness characteristics
                if (!QuickColorFilter(roi))
                {
#if DEBUG
                    try
                    {
                        Directory.CreateDirectory(_debugOutDir);
                        var dbgBmp = BitmapConverter.ToBitmap(roiMatColor);
                        var dbgPath = Path.Combine(_debugOutDir, $"tmpl_quickfilter_{DateTime.Now:yyyyMMdd_HHmmss_fff}.png");
                        dbgBmp.Save(dbgPath);
                        dbgBmp.Dispose();
                        System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] QuickColorFilter rejected ROI saved to: {dbgPath}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] QuickColorFilter debug save failed: {ex.Message}");
                    }
#endif
                    System.Diagnostics.Debug.WriteLine("[TemplateMatcher] QuickColorFilter rejected ROI - skipping template match");
                    swTotal.Stop();
                    return new MatchResult
                    {
                        Found = false,
                        Label = null,
                        BestScore = 0.0,
                        TriedCandidates = 0,
                        ElapsedMs = swTotal.ElapsedMilliseconds
                    };
                }

#if DEBUG
                try
                {
                    Directory.CreateDirectory(_debugOutDir);
                    var dbgBmp = BitmapConverter.ToBitmap(roiMatColor);
                    var dbgPath = Path.Combine(_debugOutDir, $"tmpl_roi_{DateTime.Now:yyyyMMdd_HHmmss_fff}.png");
                    dbgBmp.Save(dbgPath);
                    dbgBmp.Dispose();
                    System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] Saved ROI for debug: {dbgPath}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] ROI save failed: {ex.Message}");
                }
#endif

                // iterate templates
                foreach (var kv in _templates)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var label = kv.Key;
                    foreach (var entry in kv.Value)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        var tmpl = entry.Mat;
                        foreach (double scale in ScaleCandidates)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            int newW = (int)Math.Round(tmpl.Width * scale);
                            int newH = (int)Math.Round(tmpl.Height * scale);
                            if (newW <= 0 || newH <= 0) continue;
                            if (newW > roi.Width || newH > roi.Height) continue;

                            using var resized = new Mat();
                            Cv2.Resize(tmpl, resized, new OpenCvSharp.Size(newW, newH), 0, 0, InterpolationFlags.Lanczos4);

                            using var result = new Mat();
                            Cv2.MatchTemplate(roi, resized, result, TemplateMatchModes.CCoeffNormed);
                            Cv2.MinMaxLoc(result, out _, out double maxVal, out _, out _);

                            triedCandidates++;
                            if (maxVal > bestScore)
                            {
                                bestScore = maxVal;
                                bestLabel = label;
                                bestFile = entry.FileName;
                            }

                            // early exit if good enough
                            if (maxVal >= AcceptThreshold)
                            {
#if DEBUG
                                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] Early exit label={label} score={maxVal:0.000} >= AcceptThreshold {AcceptThreshold:0.00}. triedCandidates={triedCandidates} elapsedMs={swTotal.ElapsedMilliseconds}");
#else
                                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Early exit: score {maxVal:0.000} >= AcceptThreshold {AcceptThreshold:0.00}. triedCandidates={triedCandidates} elapsedMs={swTotal.ElapsedMilliseconds}");
#endif
                                // map label to canonical rule if possible
                                var mapped = NormalizeTemplateLabelToRule(label);
                                var outLabel = mapped ?? label;

                                // debug: log matched file
#if DEBUG
                                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] Early matched file={entry.FileName} label={label} mapped={outLabel}");
                                try
                                {
                                    Directory.CreateDirectory(_debugOutDir);
                                    var matchedTemplatePath = Path.Combine(_debugOutDir, $"tmpl_matched_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{entry.FileName}");
                                    var bmp = BitmapConverter.ToBitmap(entry.Mat);
                                    bmp.Save(matchedTemplatePath);
                                    bmp.Dispose();
                                    System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] Saved matched template: {matchedTemplatePath}");
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] Failed to save matched template: {ex.Message}");
                                }
#endif
                                swTotal.Stop();
                                return new MatchResult
                                {
                                    Found = true,
                                    Label = outLabel,
                                    BestScore = maxVal,
                                    TriedCandidates = triedCandidates,
                                    ElapsedMs = swTotal.ElapsedMilliseconds
                                };
                            }

                            // periodic cancellation check by elapsed time to avoid long-running inner loops
                            if (swTotal.ElapsedMilliseconds > MaxAllowedMsBeforeCheck)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                            }
                        } // scale
                    } // tmpl
                } // kv

                swTotal.Stop();
#if DEBUG
                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] MatchSummary elapsedMs={swTotal.ElapsedMilliseconds} triedCandidates={triedCandidates} bestScore={bestScore:0.000} bestLabel={bestLabel} bestFile={bestFile}");
#else
                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] MatchSummary elapsedMs={swTotal.ElapsedMilliseconds} triedCandidates={triedCandidates} bestScore={bestScore:0.000} bestLabel={bestLabel}");
#endif

                // final normalization of bestLabel
                var normalizedBest = NormalizeTemplateLabelToRule(bestLabel);

                // Save matched template and ROI side-by-side for inspection when found
#if DEBUG
                if (bestScore >= AcceptThreshold && !string.IsNullOrEmpty(bestFile))
                {
                    try
                    {
                        Directory.CreateDirectory(_debugOutDir);
                        foreach (var kv in _templates)
                        {
                            foreach (var entry in kv.Value)
                            {
                                if (entry.FileName == bestFile)
                                {
                                    var bmp = BitmapConverter.ToBitmap(entry.Mat);
                                    var matchedTemplatePath = Path.Combine(_debugOutDir, $"tmpl_matched_final_{DateTime.Now:yyyyMMdd_HHmmss_fff}_{bestFile}");
                                    bmp.Save(matchedTemplatePath);
                                    bmp.Dispose();
                                    System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] Saved matched template (final): {matchedTemplatePath}");
                                    goto MatchedSaved;
                                }
                            }
                        }
                    MatchedSaved:;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] Failed to save matched template (final): {ex.Message}");
                    }
                }
#endif

                return new MatchResult
                {
                    Found = bestScore >= AcceptThreshold,
                    Label = normalizedBest ?? bestLabel,
                    BestScore = bestScore == double.MinValue ? 0.0 : bestScore,
                    TriedCandidates = triedCandidates,
                    ElapsedMs = swTotal.ElapsedMilliseconds
                };
            }
            finally
            {
                if (roi != null && !roi.IsDisposed) roi.Dispose();
            }
        }, cancellationToken);
    }

    // Simple heuristic filter: check average intensity of roi; tune thresholds to your UI
    private static bool QuickColorFilter(Mat roi)
    {
        try
        {
            if (roi.Empty()) return false;
            // small downscale to speedup mean calc
            var small = new Mat();
            Cv2.Resize(roi, small, new OpenCvSharp.Size(64, Math.Max(16, Math.Min(64, roi.Height * 64 / Math.Max(1, roi.Width)))), 0, 0, InterpolationFlags.Area);
            Scalar mean = Cv2.Mean(small);
            small.Dispose();

            double avg = mean.Val0; // grayscale average 0..255
#if DEBUG
            System.Diagnostics.Debug.WriteLine($"[TemplateMatcher][DEBUG] QuickColorFilter avg={avg:0.00}");
#endif
            // adjust these thresholds to match your typical label brightness
            if (avg < 10 || avg > 245) return false;
            return true;
        }
        catch
        {
            return true; // fail-open: do not block matching on filter exceptions
        }
    }
}