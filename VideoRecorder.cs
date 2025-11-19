using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 视频录制器（使用 GDI+ 捕获屏幕）
    /// 使用 BitBlt API 捕获屏幕，简单可靠
    /// </summary>
    public sealed class VideoRecorder : IDisposable
    {
        private bool _isRecording;
        private CancellationTokenSource? _cancellationTokenSource;
        private Task? _captureTask;
        private int _frameWidth;
        private int _frameHeight;
        private int _targetWidth;
        private int _targetHeight;
        private int _frameRate;

        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hObjectSource, int nXSrc, int nYSrc, int dwRop);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        private const int SRCCOPY = 0x00CC0020;
        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        public bool IsRecording => _isRecording;

        /// <summary>
        /// 帧到达事件（每帧捕获时触发）
        /// </summary>
        public event Action<System.Drawing.Bitmap>? FrameArrived;

        /// <summary>
        /// 开始录制屏幕
        /// </summary>
        /// <param name="resolutionScale">分辨率比例 (0.1-1.0)</param>
        /// <param name="frameRate">帧率 (fps)</param>
        public void Start(double resolutionScale = 1.0, int frameRate = 30)
        {
            if (_isRecording) return;

            try
            {
                _frameRate = frameRate;
                _frameWidth = GetSystemMetrics(SM_CXSCREEN);
                _frameHeight = GetSystemMetrics(SM_CYSCREEN);
                _targetWidth = (int)(_frameWidth * resolutionScale);
                _targetHeight = (int)(_frameHeight * resolutionScale);

                _cancellationTokenSource = new CancellationTokenSource();
                _captureTask = Task.Run(() => CaptureLoop(_cancellationTokenSource.Token));
                _isRecording = true;

                WriteLine($"开始屏幕捕获: {_targetWidth}x{_targetHeight} @ {_frameRate}fps");
            }
            catch (Exception ex)
            {
                WriteError($"开始屏幕捕获失败", ex);
                throw;
            }
        }

        /// <summary>
        /// 停止录制
        /// </summary>
        public void Stop()
        {
            if (!_isRecording) return;

            try
            {
                _isRecording = false;
                _cancellationTokenSource?.Cancel();
                
                // 等待任务完成，但设置超时避免无限等待
                if (_captureTask != null)
                {
                    try
                    {
                        _captureTask.Wait(TimeSpan.FromSeconds(3)); // 3秒超时
                    }
                    catch (Exception ex)
                    {
                        WriteError($"等待捕获任务完成时出错", ex);
                    }
                }

                WriteLine("停止屏幕捕获");
            }
            catch (Exception ex)
            {
                WriteError($"停止屏幕捕获失败", ex);
            }
        }

        /// <summary>
        /// 捕获循环
        /// </summary>
        private void CaptureLoop(CancellationToken cancellationToken)
        {
            IntPtr desktopWindow = GetDesktopWindow();
            IntPtr desktopDC = GetWindowDC(desktopWindow);
            IntPtr memoryDC = CreateCompatibleDC(desktopDC);
            IntPtr bitmap = CreateCompatibleBitmap(desktopDC, _targetWidth, _targetHeight);
            IntPtr oldBitmap = SelectObject(memoryDC, bitmap);

            int frameDelay = 1000 / _frameRate; // 毫秒

            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    // 捕获屏幕
                    BitBlt(memoryDC, 0, 0, _targetWidth, _targetHeight, desktopDC, 0, 0, SRCCOPY);

                    // 转换为 Bitmap
                    using (var bmp = new Bitmap(_targetWidth, _targetHeight, PixelFormat.Format32bppArgb))
                    {
                        using (var graphics = Graphics.FromImage(bmp))
                        {
                            IntPtr hdc = graphics.GetHdc();
                            BitBlt(hdc, 0, 0, _targetWidth, _targetHeight, memoryDC, 0, 0, SRCCOPY);
                            graphics.ReleaseHdc(hdc);
                        }

                        // 触发帧到达事件（创建副本，因为 bmp 会被释放）
                        // 注意：调用者负责释放 frameCopy
                        var frameCopy = new Bitmap(bmp);
                        try
                        {
                            FrameArrived?.Invoke(frameCopy);
                        }
                        catch
                        {
                            // 如果事件处理失败，确保释放资源
                            frameCopy?.Dispose();
                            throw;
                        }
                    }

                    // 控制帧率
                    Thread.Sleep(frameDelay);
                }
            }
            finally
            {
                SelectObject(memoryDC, oldBitmap);
                DeleteObject(bitmap);
                DeleteDC(memoryDC);
                ReleaseDC(desktopWindow, desktopDC);
            }
        }

        public void Dispose()
        {
            Stop();
            _cancellationTokenSource?.Dispose();
        }
    }
}
