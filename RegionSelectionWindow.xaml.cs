using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 全屏透明窗口，用于选择录制区域
    /// </summary>
    public partial class RegionSelectionWindow : Window
    {
        private readonly Rectangle _selectionRectangle;
        private Point _startPoint;
        private bool _isSelecting;

        public Int32Rect SelectedRect { get; private set; }

        public RegionSelectionWindow()
        {
            InitializeComponent();

            // 覆盖整个虚拟屏幕，支持多显示器
            Left = SystemParameters.VirtualScreenLeft;
            Top = SystemParameters.VirtualScreenTop;
            Width = SystemParameters.VirtualScreenWidth;
            Height = SystemParameters.VirtualScreenHeight;

            _selectionRectangle = new Rectangle
            {
                Stroke = Brushes.Red,
                StrokeThickness = 2,
                Fill = new SolidColorBrush(Color.FromArgb(40, 255, 0, 0)),
                Visibility = Visibility.Collapsed
            };

            SelectionCanvas.Children.Add(_selectionRectangle);
            
            // 窗口显示时立即设置鼠标为十字光标
            this.Cursor = Cursors.Cross;
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isSelecting = true;
            CaptureMouse();
            _startPoint = e.GetPosition(this);
            Canvas.SetLeft(_selectionRectangle, _startPoint.X);
            Canvas.SetTop(_selectionRectangle, _startPoint.Y);
            _selectionRectangle.Width = 0;
            _selectionRectangle.Height = 0;
            _selectionRectangle.Visibility = Visibility.Visible;
        }

        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isSelecting) return;

            Point currentPoint = e.GetPosition(this);
            double x = Math.Min(currentPoint.X, _startPoint.X);
            double y = Math.Min(currentPoint.Y, _startPoint.Y);
            double width = Math.Abs(currentPoint.X - _startPoint.X);
            double height = Math.Abs(currentPoint.Y - _startPoint.Y);

            Canvas.SetLeft(_selectionRectangle, x);
            Canvas.SetTop(_selectionRectangle, y);
            _selectionRectangle.Width = width;
            _selectionRectangle.Height = height;
        }

        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isSelecting) return;
            _isSelecting = false;
            ReleaseMouseCapture();
            // 恢复默认鼠标样式
            this.Cursor = Cursors.Arrow;

            Rect rect = new Rect(
                Canvas.GetLeft(_selectionRectangle),
                Canvas.GetTop(_selectionRectangle),
                _selectionRectangle.Width,
                _selectionRectangle.Height);

            if (rect.Width < 10 || rect.Height < 10)
            {
                // 选择太小，视为无效
                _selectionRectangle.Visibility = Visibility.Collapsed;
                return;
            }

            try
            {
                Rect screenRect = ConvertToScreenRect(rect);
                SelectedRect = new Int32Rect(
                    (int)Math.Round(screenRect.X),
                    (int)Math.Round(screenRect.Y),
                    (int)Math.Round(screenRect.Width),
                    (int)Math.Round(screenRect.Height));

                DialogResult = true;
            }
            catch
            {
                // 转换失败，取消选择
                DialogResult = false;
            }
            finally
            {
                Cleanup();
                Close();
            }
        }

        private Rect ConvertToScreenRect(Rect rect)
        {
            // 使用 PointToScreen 获取屏幕坐标（已经是物理像素）
            // 矩形的左上角和右下角相对于窗口
            Point topLeftWindow = new Point(rect.Left, rect.Top);
            Point bottomRightWindow = new Point(rect.Right, rect.Bottom);

            // 转换为屏幕坐标（物理像素）
            Point topLeftScreen = PointToScreen(topLeftWindow);
            Point bottomRightScreen = PointToScreen(bottomRightWindow);

            // 计算物理像素尺寸
            double physicalX = topLeftScreen.X;
            double physicalY = topLeftScreen.Y;
            double physicalWidth = bottomRightScreen.X - topLeftScreen.X;
            double physicalHeight = bottomRightScreen.Y - topLeftScreen.Y;

            return new Rect(physicalX, physicalY, physicalWidth, physicalHeight);
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                // 恢复默认鼠标样式
                this.Cursor = Cursors.Arrow;
                Cleanup();
                DialogResult = false;
                Close();
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            Cleanup();
            base.OnClosed(e);
        }

        private void Cleanup()
        {
            try
            {
                if (_isSelecting)
                {
                    _isSelecting = false;
                    ReleaseMouseCapture();
                }
                // 恢复默认鼠标样式
                this.Cursor = Cursors.Arrow;
                if (_selectionRectangle != null)
                {
                    _selectionRectangle.Visibility = Visibility.Collapsed;
                }
            }
            catch
            {
                // 忽略清理错误
            }
        }
    }
}

