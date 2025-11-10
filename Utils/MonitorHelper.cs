using System;
using System.Linq;
using System.Windows;
using System.Windows.Forms;

namespace ClipDiscordApp.Utils
{
    public static class MonitorHelper
    {
        // モニタの識別子取得 (Screen.DeviceName)
        public static string GetMonitorIdContainingRect(Rect rect)
        {
            var center = new System.Drawing.Point((int)(rect.X + rect.Width / 2), (int)(rect.Y + rect.Height / 2));
            var screen = Screen.AllScreens.FirstOrDefault(s => s.Bounds.Contains(center));
            if (screen != null) return screen.DeviceName;
            // 見つからなければ最も近い画面を返す（左上との距離）
            return Screen.AllScreens.OrderBy(s => DistanceToRectCenter(s.Bounds, rect)).First().DeviceName;
        }

        private static double DistanceToRectCenter(System.Drawing.Rectangle bounds, Rect rect)
        {
            var cx = rect.X + rect.Width / 2;
            var cy = rect.Y + rect.Height / 2;
            var bx = bounds.X + bounds.Width / 2.0;
            var by = bounds.Y + bounds.Height / 2.0;
            var dx = cx - bx;
            var dy = cy - by;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        public static Screen FindScreenById(string monitorId)
        {
            if (string.IsNullOrEmpty(monitorId)) return Screen.PrimaryScreen;
            return Screen.AllScreens.FirstOrDefault(s => s.DeviceName == monitorId) ?? Screen.PrimaryScreen;
        }

        // モニタの物理Bounds（ピクセル）取得
        public static System.Drawing.Rectangle GetScreenBounds(string monitorId)
        {
            var s = FindScreenById(monitorId);
            return s.Bounds;
        }
    }
}