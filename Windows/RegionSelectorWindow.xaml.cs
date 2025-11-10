using ClipDiscordApp.Models; // DrawingRect を定義した名前空間
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ClipDiscordApp
{
    public partial class RegionSelectorWindow : Window
    {
        private System.Windows.Point? _startPoint; // 明示的に WPF の Point を使う
        private bool _isSelecting;

        public DrawingRect SelectedRegion { get; private set; } = new DrawingRect(0, 0, 0, 0);

        public RegionSelectorWindow()
        {
            InitializeComponent();
            RubberBand.Visibility = Visibility.Collapsed;
            SelectionCanvas.Cursor = System.Windows.Input.Cursors.Cross; // WPF の Cursors
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(this); // WPF の MouseButtonEventArgs を使用
            _isSelecting = true;

            Canvas.SetLeft(RubberBand, _startPoint.Value.X);
            Canvas.SetTop(RubberBand, _startPoint.Value.Y);
            RubberBand.Width = 0;
            RubberBand.Height = 0;
            RubberBand.Visibility = Visibility.Visible;

            CaptureMouse();
        }

        private void Window_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (!_isSelecting || !_startPoint.HasValue) return; // HasValue を使う
            var pos = e.GetPosition(this);

            var x = Math.Min(_startPoint.Value.X, pos.X);
            var y = Math.Min(_startPoint.Value.Y, pos.Y);
            var w = Math.Abs(pos.X - _startPoint.Value.X);
            var h = Math.Abs(pos.Y - _startPoint.Value.Y);

            Canvas.SetLeft(RubberBand, x);
            Canvas.SetTop(RubberBand, y);
            RubberBand.Width = w;
            RubberBand.Height = h;
        }

        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isSelecting || !_startPoint.HasValue) return;

            ReleaseMouseCapture();
            _isSelecting = false;

            if (RubberBand.Width < 1 || RubberBand.Height < 1)
            {
                RubberBand.Width = 1;
                RubberBand.Height = 1;
            }
        }

        private void Window_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (RubberBand.Visibility != Visibility.Visible) return;
            CommitSelectionAndClose();
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // WPF の KeyEventArgs を使う（System.Windows.Input.KeyEventArgs）
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
            else if (e.Key == Key.Enter)
            {
                if (RubberBand.Visibility == Visibility.Visible)
                {
                    CommitSelectionAndClose();
                }
            }
        }

        private void CommitSelectionAndClose()
        {
            System.Diagnostics.Debug.WriteLine("CommitSelectionAndClose called");

            var left = (int)Math.Round(Canvas.GetLeft(RubberBand));
            var top = (int)Math.Round(Canvas.GetTop(RubberBand));
            var width = (int)Math.Round(RubberBand.Width);
            var height = (int)Math.Round(RubberBand.Height);

            SelectedRegion = new DrawingRect(left, top, Math.Max(0, width), Math.Max(0, height));

            DialogResult = true;
            Close();
        }
    }
}