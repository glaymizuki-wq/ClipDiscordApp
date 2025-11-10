namespace ClipDiscordApp.Models
{
    public class AppSettings
    {
        // モニタローカル座標で保存（目的：モニタ位置が変わってもモニタ内の相対位置を保持）
        public double RegionX { get; set; }
        public double RegionY { get; set; }
        public double RegionWidth { get; set; }
        public double RegionHeight { get; set; }

        // 保存時のモニタ識別子（例: \\.\DISPLAY1）またはモニタBoundsのハッシュ
        public string MonitorId { get; set; }

        // 保存時の DPI スケール (例: 1.25)
        public double DpiScale { get; set; } = 1.0;
    }
}