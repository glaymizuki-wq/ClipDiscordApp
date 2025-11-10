using ClipDiscordApp.Models;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ClipDiscordApp.Services
{
    public class CaptureService : IDisposable
    {
        // GDI API
        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight,
            IntPtr hdcSrc, int nXSrc, int nYSrc, CopyPixelOperation dwRop);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        private bool _disposed;

        public CaptureService()
        {
        }

        /// <summary>
        /// 指定した物理ピクセル領域を同期的にキャプチャして Bitmap を返す。
        /// 呼び出しは UI スレッドで行ってもよいが重いので非同期版を推奨。
        /// </summary>
        public CaptureResult Capture(Rectangle region)
        {
            if (region.Width <= 0 || region.Height <= 0) throw new ArgumentException("region must be positive size");

            // デスクトップDC（仮想画面）を取得
            IntPtr desktopWnd = GetDesktopWindow();
            IntPtr desktopDc = GetWindowDC(desktopWnd);
            if (desktopDc == IntPtr.Zero) throw new InvalidOperationException("GetWindowDC failed");

            Bitmap bmp = null;
            IntPtr memDc = IntPtr.Zero;
            IntPtr hBitmap = IntPtr.Zero;

            try
            {
                // Create compatible DC and bitmap
                using (Graphics gDest = Graphics.FromHwnd(IntPtr.Zero))
                {
                    IntPtr hdcDest = gDest.GetHdc();
                    memDc = CreateCompatibleDC(hdcDest);
                    hBitmap = CreateCompatibleBitmap(hdcDest, region.Width, region.Height);
                    if (hBitmap == IntPtr.Zero) throw new InvalidOperationException("CreateCompatibleBitmap failed");

                    IntPtr old = SelectObject(memDc, hBitmap);

                    // BitBlt from desktopDc at region.X/Y into memDc 0/0
                    bool ok = BitBlt(memDc, 0, 0, region.Width, region.Height, desktopDc, region.X, region.Y, CopyPixelOperation.SourceCopy | CopyPixelOperation.CaptureBlt);
                    if (!ok) throw new InvalidOperationException("BitBlt failed");

                    // Create managed Bitmap from HBITMAP
                    bmp = Image.FromHbitmap(hBitmap);

                    // restore
                    SelectObject(memDc, old);
                }

                return new CaptureResult(bmp, region);
            }
            finally
            {
                // release unmanaged resources (but keep bmp for caller)
                if (hBitmap != IntPtr.Zero) DeleteObject(hBitmap);
                if (memDc != IntPtr.Zero) DeleteDC(memDc);
                if (desktopDc != IntPtr.Zero) ReleaseDC(desktopWnd, desktopDc);
            }
        }

        /// <summary>
        /// 非同期ラッパー。キャンセルをサポート。
        /// </summary>
        public Task<CaptureResult> CaptureAsync(Rectangle region, CancellationToken cancellationToken = default)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                return Capture(region);
            }, cancellationToken);
        }

        /// <summary>
        /// Bitmap を PNG として保存するユーティリティ。
        /// CaptureResult の Bitmap は呼び出し側で Dispose してください。
        /// </summary>
        public void SavePng(CaptureResult result, string path)
        {
            if (result?.Bitmap == null) throw new ArgumentNullException(nameof(result));
            result.Bitmap.Save(path, ImageFormat.Png);
        }

        #region GDI helper imports

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool DeleteDC(IntPtr hdc);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        #endregion

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }
}