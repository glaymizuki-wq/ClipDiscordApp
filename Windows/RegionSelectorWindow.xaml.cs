
using ClipDiscordApp.Models; // DrawingRect を定義した名前空間
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ClipDiscordApp
{
    public partial class RegionSelectorWindow : Window
    {
        private System.Windows.Point? _startPoint;
        private bool _isSelecting;
        private int _selectionStep = 0; // 0:文言, 1:時刻

        public DrawingRect MessageRegion { get; private set; } = new DrawingRect(0, 0, 0, 0);
        public DrawingRect TimeRegion { get; private set; } = new DrawingRect(0, 0, 0, 0);

        public RegionSelectorWindow()
        {
            InitializeComponent();
            RubberBandMessage.Visibility = Visibility.Collapsed;
            RubberBandTime.Visibility = Visibility.Collapsed;
            SelectionCanvas.Cursor = System.Windows.Input.Cursors.Cross;
            HintText.Text = "① 文言部分をドラッグして選択 → Enter/ダブルクリックで確定";
            OkButton.IsEnabled = false;
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(SelectionCanvas);
            _isSelecting = true;

            if (_selectionStep == 0)
            {
                Canvas.SetLeft(RubberBandMessage, _startPoint.Value.X);
                Canvas.SetTop(RubberBandMessage, _startPoint.Value.Y);
                RubberBandMessage.Width = 0;
                RubberBandMessage.Height = 0;
                RubberBandMessage.Visibility = Visibility.Visible;
            }
            else if (_selectionStep == 1)
            {
                Canvas.SetLeft(RubberBandTime, _startPoint.Value.X);
                Canvas.SetTop(RubberBandTime, _startPoint.Value.Y);
                RubberBandTime.Width = 0;
                RubberBandTime.Height = 0;
                RubberBandTime.Visibility = Visibility.Visible;
            }

            SelectionCanvas.CaptureMouse();
        }

        private void Window_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (!_isSelecting || !_startPoint.HasValue) return;
            var pos = e.GetPosition(SelectionCanvas);

            var x = Math.Min(_startPoint.Value.X, pos.X);
            var y = Math.Min(_startPoint.Value.Y, pos.Y);
            var w = Math.Abs(pos.X - _startPoint.Value.X);
            var h = Math.Abs(pos.Y - _startPoint.Value.Y);

            if (_selectionStep == 0)
            {
                Canvas.SetLeft(RubberBandMessage, x);
                Canvas.SetTop(RubberBandMessage, y);
                RubberBandMessage.Width = w;
                RubberBandMessage.Height = h;
            }
            else if (_selectionStep == 1)
            {
                Canvas.SetLeft(RubberBandTime, x);
                Canvas.SetTop(RubberBandTime, y);
                RubberBandTime.Width = w;
                RubberBandTime.Height = h;
            }
        }

        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isSelecting || !_startPoint.HasValue) return;

            SelectionCanvas.ReleaseMouseCapture();
            _isSelecting = false;

            if (_selectionStep == 0 && (RubberBandMessage.Width < 1 || RubberBandMessage.Height < 1))
            {
                RubberBandMessage.Width = 1;
                RubberBandMessage.Height = 1;
            }
            else if (_selectionStep == 1 && (RubberBandTime.Width < 1 || RubberBandTime.Height < 1))
            {
                RubberBandTime.Width = 1;
                RubberBandTime.Height = 1;
            }
        }

        private void Window_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            CommitSelectionStep();
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
            else if (e.Key == Key.Enter)
            {
                CommitSelectionStep();
            }
        }

        private void CommitSelectionStep()
        {
            if (_selectionStep == 0)
            {
                // 文言部分の領域確定
                var left = (int)Math.Round(Canvas.GetLeft(RubberBandMessage));
                var top = (int)Math.Round(Canvas.GetTop(RubberBandMessage));
                var width = (int)Math.Round(RubberBandMessage.Width);
                var height = (int)Math.Round(RubberBandMessage.Height);
                MessageRegion = new DrawingRect(left, top, Math.Max(0, width), Math.Max(0, height));

                _selectionStep = 1;
                HintText.Text = "② 時刻部分をドラッグして選択 → Enter/ダブルクリックで確定";
            }
            else if (_selectionStep == 1)
            {
                // 時刻部分の領域確定
                var left = (int)Math.Round(Canvas.GetLeft(RubberBandTime));
                var top = (int)Math.Round(Canvas.GetTop(RubberBandTime));
                var width = (int)Math.Round(RubberBandTime.Width);
                var height = (int)Math.Round(RubberBandTime.Height);
                TimeRegion = new DrawingRect(left, top, Math.Max(0, width), Math.Max(0, height));

                OkButton.IsEnabled = true;
                HintText.Text = "選択完了！OKボタンで確定できます";
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            _selectionStep = 0;
            RubberBandMessage.Visibility = Visibility.Collapsed;
            RubberBandTime.Visibility = Visibility.Collapsed;
            HintText.Text = "① 文言部分をドラッグして選択 → Enter/ダブルクリックで確定";
            OkButton.IsEnabled = false;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageRegion.Width > 0 && MessageRegion.Height > 0 &&
                TimeRegion.Width > 0 && TimeRegion.Height > 0)
            {
                DialogResult = true;
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}