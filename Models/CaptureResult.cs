using System;
using System.Drawing;

namespace ClipDiscordApp.Models
{
    // キャプチャ結果（Bitmap を所有し、領域情報を保持）
    public class CaptureResult : IDisposable
    {
        public Bitmap Bitmap { get; private set; }
        public Rectangle Region { get; private set; }      // キャプチャに使った領域（物理ピクセル）
        public DateTime Timestamp { get; private set; }    // キャプチャ時刻（任意）

        private bool _disposed;

        public CaptureResult(Bitmap bitmap, Rectangle region)
        {
            Bitmap = bitmap ?? throw new ArgumentNullException(nameof(bitmap));
            Region = region;
            Timestamp = DateTime.UtcNow;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;
            if (disposing)
            {
                Bitmap?.Dispose();
            }
            _disposed = true;
        }

        ~CaptureResult()
        {
            Dispose(false);
        }
    }
}