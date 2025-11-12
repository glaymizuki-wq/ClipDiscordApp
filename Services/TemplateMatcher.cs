using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace ClipDiscordApp.Services
{
    public static class TemplateMatcher
    {
        private static readonly Dictionary<string, List<Mat>> _templates = new();

        // 呼び出し: TemplateMatcher.Initialize(templatesRootPath)
        public static void Initialize(string templatesRoot)
        {
            _templates.Clear();
            if (!Directory.Exists(templatesRoot)) return;

            foreach (var labelDir in Directory.GetDirectories(templatesRoot))
            {
                var label = Path.GetFileName(labelDir).ToUpperInvariant();
                var list = new List<Mat>();
                foreach (var f in Directory.GetFiles(labelDir).Where(x => IsImageFile(x)))
                {
                    try
                    {
                        using var bmp = (Bitmap)System.Drawing.Image.FromFile(f);
                        var mat = BitmapConverter.ToMat(bmp);
                        Cv2.CvtColor(mat, mat, ColorConversionCodes.BGR2GRAY);
                        mat.ConvertTo(mat, MatType.CV_32F);
                        Cv2.Normalize(mat, mat, 0, 1, NormTypes.MinMax);
                        list.Add(mat);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] load template failed {f}: {ex.Message}");
                    }
                }
                if (list.Count > 0) _templates[label] = list;
            }

            var summary = string.Join(", ", _templates.Select(kv => $"{kv.Key}:{kv.Value.Count}"));
            System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Initialized templates: {summary}");
        }

        // bitmap: 前処理済み(prepped) の Bitmap を渡す
        public static (bool found, string label, double score) Check(Bitmap bitmap, double minScore = 0.75, double scaleMin = 0.9, double scaleMax = 1.1, double scaleStep = 0.05)
        {
            if (bitmap == null) return (false, null, 0.0);
            if (_templates.Count == 0) return (false, null, 0.0);

            Mat src = null;
            try
            {
                src = BitmapConverter.ToMat(bitmap);
                Cv2.CvtColor(src, src, ColorConversionCodes.BGR2GRAY);
                src.ConvertTo(src, MatType.CV_32F);
                Cv2.Normalize(src, src, 0, 1, NormTypes.MinMax);

                string bestLabel = null;
                double bestScore = 0.0;

                foreach (var kv in _templates)
                {
                    var label = kv.Key;
                    foreach (var tmpl in kv.Value)
                    {
                        for (double scale = scaleMin; scale <= scaleMax; scale += scaleStep)
                        {
                            var scaledW = (int)(tmpl.Width * scale);
                            var scaledH = (int)(tmpl.Height * scale);
                            if (scaledW < 3 || scaledH < 3) continue;
                            if (scaledW > src.Width || scaledH > src.Height) continue;

                            using var tmplResized = new Mat();
                            try
                            {
                                Cv2.Resize(tmpl, tmplResized, new OpenCvSharp.Size(scaledW, scaledH), 0, 0, InterpolationFlags.Linear);
                                var resultCols = src.Width - tmplResized.Width + 1;
                                var resultRows = src.Height - tmplResized.Height + 1;
                                if (resultCols <= 0 || resultRows <= 0) continue;

                                using var result = new Mat(resultRows, resultCols, MatType.CV_32F);
                                Cv2.MatchTemplate(src, tmplResized, result, TemplateMatchModes.CCoeffNormed);
                                Cv2.MinMaxLoc(result, out _, out double maxVal, out _, out _);

                                if (maxVal > bestScore)
                                {
                                    bestScore = maxVal;
                                    bestLabel = label;
                                }
                            }
                            catch { /* ignore */ }
                        }
                    }
                }

                if (bestScore >= minScore && !string.IsNullOrEmpty(bestLabel))
                {
                    return (true, bestLabel, bestScore);
                }

                return (false, null, bestScore);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TemplateMatcher] Check error: {ex.Message}");
                return (false, null, 0.0);
            }
            finally
            {
                src?.Dispose();
            }
        }

        private static bool IsImageFile(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".bmp";
        }
    }
}