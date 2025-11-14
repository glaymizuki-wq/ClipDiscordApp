using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using Tesseract; // 必須パッケージ参照

namespace ClipDiscordApp.Utils
{
    public static class OcrTextUtils
    {
        // existing behaviour
        public static int FuzzyThreshold { get; set; } = 2;
        public static int GetFuzzyThreshold() => FuzzyThreshold;

        public static string NormalizeForComparison(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            s = s.ToUpperInvariant();

            var sb = new System.Text.StringBuilder();
            foreach (var ch in s)
            {
                if (char.IsLetterOrDigit(ch) || ch == ':' || ch == '-' || ch == ' ') sb.Append(ch);
            }
            s = sb.ToString();

            s = s.Replace('0', 'O');   // 0 -> O
            s = s.Replace('1', 'I');   // 1 -> I
            s = s.Replace('5', 'S');   // 5 -> S
            s = s.Replace('8', 'B');   // 8 -> B
            s = s.Replace('|', 'I');   // | -> I
            s = s.Replace("::", ":");  // 重複コロン修正
            s = s.Replace(" ", "");    // 比較用は空白除去

            // TEST mappings (keep if useful)
            s = s.Replace("SML", "SELL");
            s = s.Replace("SMI", "SELL");
            s = s.Replace("SEI", "SELL");
            s = s.Replace("S M L", "SELL");
            s = s.Replace("0SE", "OSE");
            return s;
        }

        public static string NormalizeForOutput(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            return s.Trim();
        }

        public static void LogFuzzyMatch(string token, string pattern, int dist)
        {
            System.Diagnostics.Debug.WriteLine($"[Fuzzy] token='{token}' pattern='{pattern}' dist={dist}");
        }

        public static int LevenshteinDistance(string a, string b)
        {
            if (a == null) a = "";
            if (b == null) b = "";
            int n = a.Length, m = b.Length;
            if (n == 0) return m;
            if (m == 0) return n;
            var d = new int[n + 1, m + 1];
            for (int i = 0; i <= n; i++) d[i, 0] = i;
            for (int j = 0; j <= m; j++) d[0, j] = j;
            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }
            return d[n, m];
        }

        // ----------------- ここから追加ユーティリティ -----------------

        // Shared TesseractEngine
        private static TesseractEngine _sharedEngine;
        private static readonly object _engineLock = new();

        // 既存の GetTesseractEngine 呼び出しを置き換えるための公開メソッド
        public static TesseractEngine GetEngine(string tessdataRelativePath = "tessdata", string lang = "eng")
        {
            lock (_engineLock)
            {
                if (_sharedEngine != null) return _sharedEngine;

                var exeDir = AppDomain.CurrentDomain.BaseDirectory;
                var tessDataPath = Path.Combine(exeDir, tessdataRelativePath);
                if (!Directory.Exists(tessDataPath))
                {
                    Debug.WriteLine($"[OcrTextUtils] tessdata not found at {tessDataPath}");
                    throw new DirectoryNotFoundException($"tessdata not found: {tessDataPath}");
                }

                try
                {
                    _sharedEngine = new TesseractEngine(tessDataPath, lang, EngineMode.Default);
                    Debug.WriteLine("[OcrTextUtils] TesseractEngine created");
                    return _sharedEngine;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[OcrTextUtils] Failed to create TesseractEngine: {ex}");
                    throw;
                }
            }
        }

        public static void DisposeEngine()
        {
            lock (_engineLock)
            {
                try { _sharedEngine?.Dispose(); } catch { }
                _sharedEngine = null;
            }
        }

        // Bitmap -> Pix 変換。呼び出し側が Dispose を行うこと（using を推奨）
        public static Pix BitmapToPix(Bitmap bmp)
        {
            if (bmp == null) return null;

            // Try to use PixConverter if available
            try
            {
                var convType = Type.GetType("Tesseract.PixConverter, Tesseract");
                if (convType != null)
                {
                    var toPix = convType.GetMethod("ToPix", new Type[] { typeof(Bitmap) });
                    if (toPix != null)
                    {
                        return (Pix)toPix.Invoke(null, new object[] { bmp });
                    }
                }
            }
            catch
            {
                // ignore and fallback
            }

            // Fallback: attempt direct conversion (best-effort)
            try
            {
                using var bmp32 = new Bitmap(bmp.Width, bmp.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(bmp32)) g.DrawImage(bmp, 0, 0, bmp.Width, bmp.Height);

                var data = bmp32.LockBits(new Rectangle(0, 0, bmp32.Width, bmp32.Height),
                    System.Drawing.Imaging.ImageLockMode.ReadOnly, bmp32.PixelFormat);
                try
                {
                    int stride = Math.Abs(data.Stride);
                    int size = stride * bmp32.Height;
                    var bytes = new byte[size];
                    System.Runtime.InteropServices.Marshal.Copy(data.Scan0, bytes, 0, size);

                    var pixType = typeof(Pix);
                    var loadMethod = pixType.GetMethod("LoadFromMemory", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                    if (loadMethod != null)
                    {
                        var pix = (Pix)loadMethod.Invoke(null, new object[] { bytes, bytes.Length });
                        if (pix != null) return pix;
                    }
                }
                finally
                {
                    bmp32.UnlockBits(data);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OcrTextUtils] BitmapToPix fallback failed: {ex.Message}");
            }

            throw new NotSupportedException("Bitmap -> Pix conversion not supported in this environment. Add Pix converter.");
        }

        // LabelCandidate のスコアを安全に取り出す（reflection）
        public static double GetCandidateScore(object candidate)
        {
            if (candidate == null) return 0.0;
            var t = candidate.GetType();
            string[] names = new[] { "Score", "score", "Confidence", "confidence", "Value", "value" };
            foreach (var n in names)
            {
                var p = t.GetProperty(n);
                if (p != null && IsNumericType(p.PropertyType))
                {
                    var val = p.GetValue(candidate);
                    return Convert.ToDouble(val);
                }
                var f = t.GetField(n);
                if (f != null && IsNumericType(f.FieldType))
                {
                    var val = f.GetValue(candidate);
                    return Convert.ToDouble(val);
                }
            }
            var numericProp = t.GetProperties().FirstOrDefault(pp => IsNumericType(pp.PropertyType));
            if (numericProp != null) return Convert.ToDouble(numericProp.GetValue(candidate));
            return 0.0;
        }

        private static bool IsNumericType(Type t)
        {
            return t == typeof(double) || t == typeof(float) || t == typeof(decimal)
                || t == typeof(int) || t == typeof(long) || t == typeof(short);
        }
    }
}