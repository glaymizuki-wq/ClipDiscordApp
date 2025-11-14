// TemplateMatcher.cs (modified)
// Requires: OpenCvSharp, OpenCvSharp.Extensions
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OpenCvSharp;
using OpenCvSharp.Extensions;

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
    // label -> list of template Mats (grayscale, preprocessed)
    private static readonly Dictionary<string, List<Mat>> _templates = new();

    // Configurable options
    public static double AcceptThreshold { get; set; } = 0.80; // default threshold (tuneable)
    public static double[] ScaleCandidates { get; set; } = new[] { 0.98, 1.0, 1.02 }; // reduce candidates for speed
    public static int MaxAllowedMsBeforeCheck { get; set; } = 500; // check cancellation every N ms

    // Initialize: load templates from folder (file names should include label as second segment e.g. tpl_buy_01.png)
    public static void Initialize(string templatesDir)
    {
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
                    list = new List<Mat>();
                    _templates[label] = list;
                }
                list.Add(gray);
                loadedFiles.Add(Path.GetFileName(f));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Load failed {f}: {ex.Message}");
            }
        }

        System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Loaded labels: {_templates.Keys.Count}");
        if (loadedFiles.Count > 0)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Found files: {string.Join(", ", loadedFiles)}");
        }
    }

    // Async Check that respects cancellation and logs summary
    // preppedBmp must be preprocessed with the same pipeline used to create templates (grayscale/resized)
    public static async Task<MatchResult> CheckAsync(System.Drawing.Bitmap preppedBmp, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() =>
        {
            var swTotal = System.Diagnostics.Stopwatch.StartNew();
            int triedCandidates = 0;
            double bestScore = double.MinValue;
            string bestLabel = null;

            if (_templates.Count == 0)
            {
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

                // iterate templates with cancellation checks and early exit
                foreach (var kv in _templates)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var label = kv.Key;
                    foreach (var tmpl in kv.Value)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        foreach (double scale in ScaleCandidates)
                        {
                            cancellationToken.ThrowIfCancellationRequested();

                            // scale template
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
                            }

                            // early exit if good enough
                            if (maxVal >= AcceptThreshold)
                            {
                                swTotal.Stop();
                                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Early exit: score {maxVal:0.000} >= AcceptThreshold {AcceptThreshold:0.00}. triedCandidates={triedCandidates} elapsedMs={swTotal.ElapsedMilliseconds}");
                                return new MatchResult
                                {
                                    Found = true,
                                    Label = label,
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
                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] MatchSummary elapsedMs={swTotal.ElapsedMilliseconds} triedCandidates={triedCandidates} bestScore={bestScore:0.000} bestLabel={bestLabel}");
                return new MatchResult
                {
                    Found = bestScore >= AcceptThreshold,
                    Label = bestLabel,
                    BestScore = bestScore == double.MinValue ? 0.0 : bestScore,
                    TriedCandidates = triedCandidates,
                    ElapsedMs = swTotal.ElapsedMilliseconds
                };
            }
            finally
            {
                roi.Dispose();
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