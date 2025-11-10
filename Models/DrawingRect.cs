namespace ClipDiscordApp.Models
{
    // 軽量な矩形クラス（System.Drawing.Rectangle へ変換して使います）
    public class DrawingRect
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        public DrawingRect() { }
        public DrawingRect(int x, int y, int w, int h)
        {
            X = x; Y = y; Width = w; Height = h;
        }
    }
}