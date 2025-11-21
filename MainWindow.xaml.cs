using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Threading;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    public partial class MainWindow : Window
    {
        private readonly AudioRecorder _audioRecorder = new();
        private VideoEncoder? _videoEncoder;
        private RecordingConfig _config;
        private readonly string _workDir;
        private readonly string _configPath;
        private bool _isVideoRecording;
        private RegionHighlightWindow? _regionHighlightWindow;
        
        // 截图相关
        private DispatcherTimer? _screenshotTimer;
        private Bitmap? _lastScreenshot;
        private DateTime _lastScreenshotTime;
        private bool _isScreenshotEnabled;
        private readonly string _screenshotDir;
        
        // PPT和PDF生成器
        private PPTGenerator? _pptGenerator;
        private PDFGenerator? _pdfGenerator;
        private string? _pptFilePath;
        private string? _pdfFilePath;
        private bool _pptPdfInitialized = false;
        
        // 录制状态信息
        private DateTime _recordingStartTime;
        private int _screenshotCount = 0;
        private DispatcherTimer? _statusUpdateTimer;
        private string _currentOperation = "";
        private double _currentScreenChangeRate = 0.0; // 实时检测到的屏幕变化率
        private bool _isStoppingRecording = false;

        public MainWindow()
        {
            InitializeComponent();
            
            // 设置窗口始终置顶
            this.Topmost = true;
            
            // 设置窗口初始位置在屏幕顶部居中
            this.WindowStartupLocation = WindowStartupLocation.Manual;
            this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
            this.Top = 10; // 距离顶部10像素
            
            // 添加窗口拖拽功能（无边框窗口需要手动实现）
            this.MouseDown += MainWindow_MouseDown;
            
            // 设置窗口关闭事件
            this.Closing += MainWindow_Closing;
            
            // 获取 .exe 所在目录
            string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
            
            // 配置文件路径：放在 .exe 所在目录
            _configPath = Path.Combine(exeDirectory, "config.json");

            // 设置工作目录：优先使用D盘，如果D盘不存在则使用C盘（用于保存录制的视频文件）
            string driveLetter = Directory.Exists("D:\\") ? "D:" : "C:";
            _workDir = Path.Combine(driveLetter, "ScreenshotV3.0");
            Directory.CreateDirectory(_workDir);
            
            // 设置截图目录：优先使用D盘，如果D盘不存在则使用C盘
            _screenshotDir = Path.Combine(driveLetter, "ScreenshotV3.0");
            Directory.CreateDirectory(_screenshotDir);

            // 加载配置
            _config = RecordingConfig.Load(_configPath);
            
            // 初始化日志系统
            Logger.Enabled = _config.LogEnabled == 1;
            Logger.SetLogFileMode(_config.LogFileMode);
            // 设置日志文件目录为工作目录（与视频、音频文件同一目录）
            // SetLogDirectory 会根据模式决定是否清空日志文件
            Logger.SetLogDirectory(_workDir);
            
            // 初始化截图定时器（但不启动，只有点击开始按钮后才启动）
            _screenshotTimer = new DispatcherTimer();
            _screenshotTimer.Tick += ScreenshotTimer_Tick;
            _lastScreenshotTime = DateTime.Now;
            
            // 初始化截图功能（默认开启，但定时器不启动）
            _isScreenshotEnabled = true;
            if (MenuScreenshot != null)
            {
                MenuScreenshot.IsChecked = true;
            }
            
            // 注意：截图定时器不在启动时自动运行，只有点击"开始"按钮后才启动
            
            // 初始化菜单项状态
            UpdateLogMenuItems();
            UpdatePPTAndPDFMenuItems();
            
            // 初始化界面上的设置输入框
            InitializeSettingsControls();
            
            Logger.WriteLine("程序启动");
            Logger.WriteLine($"日志状态: {(Logger.Enabled ? "启用" : "禁用")}");
            Logger.WriteLine($"日志文件模式: {(_config.LogFileMode == 0 ? "覆盖" : "叠加")}");
            
            UpdateConfigDisplay();
            
            // 初始化按钮状态：只有选择按钮可用，开始和停止按钮禁用
            InitializeButtonStates();
            
            // 延迟初始化区域高亮和截图基准画面，直到窗口加载完成
            this.Loaded += (s, e) =>
            {
                // 窗口加载完成后，初始化 _lastScreenshot 作为第一次检查的基准画面
                if (_isScreenshotEnabled)
                {
                    UpdateLastScreenshot();
                }
            };
        }

        private void MainWindow_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // 实现无边框窗口拖拽功能（只在非按钮区域才能拖拽）
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left && 
                e.Source is not System.Windows.Controls.Button)
            {
                this.DragMove();
            }
        }

        private void BtnMinimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // 立即允许关闭，不阻塞
            e.Cancel = false;

            Dispatcher.Invoke(() =>
            {
                _regionHighlightWindow?.Close();
                _regionHighlightWindow = null;
            });

            // 在后台线程执行清理，不阻塞窗口关闭
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    // 如果正在录制，快速停止（不等待完成）
                    if (_isVideoRecording)
                    {
                        try
                        {
                            // 快速停止，不等待完成（FFmpeg 会自动停止音频录制）
                            // 使用快速退出模式完成编码，立即终止 FFmpeg
                            if (_videoEncoder != null)
                            {
                                try
                                {
                                    _videoEncoder.Finish(quickExit: true);
                                }
                                catch (Exception ex)
                                {
                                    WriteError($"完成编码失败", ex);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteError($"停止录制时出错", ex);
                        }
                    }
                    else
                    {
                        // 如果只是音频录制，也停止
                        try
                        {
                            _audioRecorder.Stop();
                        }
                        catch (Exception ex)
                        {
                            WriteError($"停止音频录制时出错", ex);
                        }
                    }

                    // 延迟释放资源，给 Finish 一些时间
                    Thread.Sleep(1000);
                    
                    // 释放资源
                    try
                    {
                        _videoEncoder?.Dispose();
                        _audioRecorder?.Dispose();
                    }
                    catch (Exception ex)
                    {
                        WriteError($"释放资源时出错", ex);
                    }
                    
                    // 停止截图定时器
                    try
                    {
                        if (_screenshotTimer != null && _screenshotTimer.IsEnabled)
                        {
                            _screenshotTimer.Stop();
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteError($"停止截图定时器时出错", ex);
                    }
                }
                catch (Exception ex)
                {
                    WriteError($"窗口关闭处理时出错", ex);
                }
            });
        }

        private void MenuSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var settingsWindow = new SettingsWindow(_config);
                if (settingsWindow.ShowDialog() == true)
                {
                    _config = settingsWindow.Config;
                    _config.Save(_configPath);
                    UpdateConfigDisplay();
                    UpdateLogMenuItems();
                    UpdatePPTAndPDFMenuItems();
                    UpdateStatusDisplay("设置已保存");
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"打开设置失败：{ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMsg += $"\n详细：{ex.InnerException.Message}";
                }
                UpdateStatusDisplay(errorMsg);
                WriteError($"打开设置窗口异常", ex);
            }
        }

        private void BtnStartVideo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isVideoRecording) return;
                _isStoppingRecording = false;

                // 每次开始录制时，清理之前的PPT/PDF生成器，确保生成新的文件
                // 因为用户可能重新选择了区域，需要生成新的文件
                if (_pptGenerator != null || _pdfGenerator != null)
                {
                    WriteLine("清理之前的PPT/PDF生成器，准备生成新文件");
                    _pptGenerator?.Dispose();
                    _pdfGenerator?.Dispose();
                    _pptGenerator = null;
                    _pdfGenerator = null;
                    _pptPdfInitialized = false;
                }

                var videoPath = Path.Combine(_workDir, $"video{DateTime.Now:yyMMddHHmmss}.mp4");

                int videoWidth;
                int videoHeight;
                int offsetX = 0;
                int offsetY = 0;

                var (screenWidthPixels, screenHeightPixels) = GetPrimaryScreenPixelSize();

                if (HasValidCustomRegion())
                {
                    int regionRight = _config.RegionLeft + _config.RegionWidth;
                    int regionBottom = _config.RegionTop + _config.RegionHeight;
                    
                    if (_config.RegionLeft < 0 || _config.RegionTop < 0 ||
                        regionRight > screenWidthPixels || regionBottom > screenHeightPixels)
                    {
                        string errorMsg = $"选择的录制区域超出主屏幕范围！\n" +
                                         $"区域: ({_config.RegionLeft}, {_config.RegionTop}) 到 ({regionRight}, {regionBottom})\n" +
                                         $"主屏幕: (0, 0) 到 ({screenWidthPixels}, {screenHeightPixels})\n" +
                                         $"请重新选择区域或清除区域设置使用全屏录制。";
                        
                        WriteError(errorMsg);
                        UpdateStatusDisplay("录制区域超出屏幕范围，请重新选择");
                        
                        MessageBox.Show(errorMsg, "区域超出范围", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    
                    videoWidth = _config.RegionWidth;
                    videoHeight = _config.RegionHeight;
                    offsetX = _config.RegionLeft;
                    offsetY = _config.RegionTop;
                    WriteLine($"使用自定义区域录制: {videoWidth}x{videoHeight} @ 起点 ({offsetX}, {offsetY})");
                }
                else
                {
                    double resolutionScale = _config.VideoResolutionScale / 100.0;
                    videoWidth = (int)(screenWidthPixels * resolutionScale);
                    videoHeight = (int)(screenHeightPixels * resolutionScale);
                }

                // 创建视频编码器（不再需要 VideoRecorder，FFmpeg 直接录制）
                _videoEncoder = new VideoEncoder(videoPath, _config);

                // 初始化编码器（支持自定义区域偏移）
                _videoEncoder.Initialize(videoWidth, videoHeight, _config.VideoFrameRate, 
                    _config.AudioSampleRate, 2, offsetX, offsetY);

                // 开始音频录制（使用 NAudio WasapiLoopbackCapture，无需虚拟声卡）
                // 注意：无论是否有声音输出，都要录制音频（包括静音）
                // AudioRecorder.Start() 会创建与目标同名的临时 wav 文件
                var timestamp = DateTime.Now.ToString("yyMMddHHmmss");
                var audioOutputPath = Path.Combine(_workDir, $"audio_{timestamp}.m4a"); // 最终输出路径（不会被使用）
                var expectedAudioPath = Path.Combine(_workDir, $"audio_{timestamp}.wav"); // AudioRecorder 实际创建的文件
                WriteLine($"========== 开始录制 ==========");
                WriteLine($"视频文件: {videoPath}");
                WriteLine($"开始音频录制（NAudio WasapiLoopbackCapture，即使静音也会录制），输出路径: {audioOutputPath}");
                WriteLine($"预期音频文件: {expectedAudioPath}");
                
                // 连接音频数据事件，实现边录边合成
                _audioRecorder.AudioSampleAvailable += OnAudioSampleAvailable;
                
                _audioRecorder.Start(audioOutputPath);
                
                // 等待音频录制启动
                System.Threading.Thread.Sleep(300);
                
                // 启动 FFmpeg 录制视频
                // VideoEncoder.Start() 会尝试直接录制有声视频，如果失败则使用音频管道实时合成
                WriteLine($"启动 FFmpeg 录制视频...");
                _videoEncoder.Start();

                _isVideoRecording = true;
                // 更新按钮状态
                UpdateStartButtonState();
                
                // 初始化录制状态信息
                _recordingStartTime = DateTime.Now;
                _screenshotCount = 0;
                
                // 启动状态更新定时器
                _statusUpdateTimer = new DispatcherTimer();
                _statusUpdateTimer.Interval = TimeSpan.FromMilliseconds(500); // 每500ms更新一次
                _statusUpdateTimer.Tick += StatusUpdateTimer_Tick;
                _statusUpdateTimer.Start();

                // 启动截图定时器（只有在开始录制时才启动）
                if (_isScreenshotEnabled && _screenshotTimer != null)
                {
                    _screenshotTimer.Interval = TimeSpan.FromSeconds(1);
                    _screenshotTimer.Start();
                    _lastScreenshotTime = DateTime.Now;
                    // 立即初始化 _lastScreenshot，作为第一次检查的基准画面
                    UpdateLastScreenshot();
                    WriteLine("截图定时器已启动");
                }

                // 状态显示会在定时器中更新，显示：屏幕变化率、录制间隔、截图数量
                // 立即更新一次状态
                UpdateRecordingStatus();
                
                // 如果截图功能开启，立即截图一张
                if (_isScreenshotEnabled)
                {
                    try
                    {
                        CaptureScreenshot();
                        WriteLine("开始录制时立即截图一张");
                    }
                    catch (Exception ex)
                    {
                        WriteError($"开始录制时截图失败", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatusDisplay($"开始录制视频失败：{ex.Message}");
            }
        }

        private void BtnStopVideo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!_isVideoRecording) return;

                // 先禁用按钮，防止重复点击
                BtnStopVideo.IsEnabled = false;
                
                _isStoppingRecording = true;
                _statusUpdateTimer?.Stop();
                _statusUpdateTimer = null;
                
                // 立即更新状态显示
                _currentOperation = "正在停止录制...";
                UpdateStatusDisplayWithScroll(_currentOperation, System.Windows.Media.Brushes.Orange);

                // 在后台线程执行停止操作，避免阻塞 UI
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    try
                    {
                        // 同时停止视频和音频录制，确保时长一致
                        WriteLine($"========== 停止录制 ==========");
                        WriteLine($"同时停止视频和音频录制...");
                        
                        // 先发送停止信号给 FFmpeg（不等待完成）
                        _videoEncoder?.RequestStop();
                        
                        // 断开音频事件连接（避免继续写入数据）
                        _audioRecorder.AudioSampleAvailable -= OnAudioSampleAvailable;
                        
                        // 立即停止音频录制（与视频同时停止）
                        WriteLine($"停止音频录制");
                        _audioRecorder.Stop();
                        
                        // 更新状态：正在生成MP4文件
                        Dispatcher.Invoke(() =>
                        {
                            _currentOperation = "正在生成MP4文件...";
                            UpdateStatusDisplayWithScroll(_currentOperation, System.Windows.Media.Brushes.Orange);
                        });
                        
                        // 等待音频管道数据刷新（实时合成模式）或音频文件写入完成（文件合并模式）
                        System.Threading.Thread.Sleep(1000);
                        
                        // 检查是否是实时合成模式（_hasAudioInVideo = true 且没有找到音频设备）
                        // 如果是实时合成模式，不需要查找和合并音频文件
                        // 如果是文件合并模式，需要查找音频文件并合并
                        
                        // 注意：VideoEncoder.Finish() 会根据 _hasAudioInVideo 判断是否需要合并
                        // 如果是实时合成模式，_hasAudioInVideo = true，Finish() 会跳过合并
                        // 如果是文件合并模式，_hasAudioInVideo = false，Finish() 会查找 _tempAudioPath 并合并
                        
                        // 只有在文件合并模式下，才需要查找和设置音频文件路径
                        // 实时合成模式下，音频已经通过管道写入 FFmpeg，无需合并
                        
                        // 尝试查找临时音频文件（仅在文件合并模式下需要）
                        var tempAudioFiles = Directory.GetFiles(_workDir, "audio_*.wav");
                        if (tempAudioFiles.Length == 0)
                        {
                            tempAudioFiles = Directory.GetFiles(_workDir, "temp_audio_*.wav");
                        }
                        
                        if (tempAudioFiles.Length > 0)
                        {
                            var latestAudioFile = tempAudioFiles.OrderByDescending(f => File.GetCreationTime(f)).First();
                            var audioFileInfo = new FileInfo(latestAudioFile);
                            
                            // 设置音频文件路径（VideoEncoder.Finish() 会根据 _hasAudioInVideo 决定是否使用）
                            if (audioFileInfo.Length > 0)
                            {
                                _videoEncoder?.SetAudioFile(latestAudioFile);
                            }
                        }
                        
                        // 现在停止视频录制并完成编码
                        // 注意：如果是实时合成模式，Finish() 会跳过合并；如果是文件合并模式，Finish() 会合并音频
                        WriteLine($"准备停止视频录制（FFmpeg）并完成编码");
                        _videoEncoder?.Finish(); // 这会发送 'q' 给 FFmpeg，等待退出，然后根据模式决定是否合并音频
                        WriteLine($"视频编码完成");

                        // 释放资源
                        _videoEncoder?.Dispose();

                        // 完成PPT和PDF生成（在MP4生成完成后立即执行）
                        if (_config.GeneratePPT || _config.GeneratePDF)
                        {
                            if (_config.GeneratePPT)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    _currentOperation = "正在生成PPT文件...";
                                    UpdateStatusDisplayWithScroll(_currentOperation, System.Windows.Media.Brushes.Orange);
                                });
                            }
                            
                            if (_config.GeneratePDF)
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    _currentOperation = "正在生成PDF文件...";
                                    UpdateStatusDisplayWithScroll(_currentOperation, System.Windows.Media.Brushes.Orange);
                                });
                            }
                        }
                        
                        WriteLine($"准备完成PPT和PDF生成");
                        FinalizePPTAndPDF();
                        WriteLine($"PPT和PDF生成完成");

                        // 在 UI 线程更新界面
                        Dispatcher.Invoke(() =>
                        {
                            _videoEncoder = null;
                            _isVideoRecording = false;
                            
                            // 停止截图定时器
                            if (_screenshotTimer != null && _screenshotTimer.IsEnabled)
                            {
                                _screenshotTimer.Stop();
                                WriteLine("截图定时器已停止");
                            }
                            
                            // 更新按钮状态（考虑区域选择情况）
                            UpdateStartButtonState();
                            
                            // 清除当前操作，恢复正常状态显示
                            _isStoppingRecording = false;
                            _currentOperation = "";
                            UpdateStatusDisplay("视频录制已停止");
                        });
                    }
                    catch (Exception ex)
                    {
                        WriteError($"停止录制失败", ex);
                        Dispatcher.Invoke(() =>
                        {
                            // 停止状态更新定时器
                            _statusUpdateTimer?.Stop();
                            _statusUpdateTimer = null;
                            _currentOperation = "";
                            _isStoppingRecording = false;
                            
                            UpdateStatusDisplay($"停止录制失败：{ex.Message}");
                            UpdateStartButtonState();
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                // 停止状态更新定时器
                _statusUpdateTimer?.Stop();
                _statusUpdateTimer = null;
                _currentOperation = "";
                _isStoppingRecording = false;
                
                UpdateStatusDisplay($"停止录制视频失败：{ex.Message}");
                UpdateStartButtonState();
            }
        }

        private void OnAudioSampleAvailable(byte[]? audioData, int bytesRecorded)
        {
            // 音频样本可用事件：将音频数据实时传递给 FFmpeg（边录边合成）
            if (audioData != null && bytesRecorded > 0 && _videoEncoder != null)
            {
                _videoEncoder.WriteAudioData(audioData, bytesRecorded);
            }
        }


        private void UpdateConfigDisplay()
        {
            UpdateStatusDisplay(null);
        }

        /// <summary>
        /// 状态更新定时器事件
        /// </summary>
        private void StatusUpdateTimer_Tick(object? sender, EventArgs e)
        {
            if (_isVideoRecording && !_isStoppingRecording)
            {
                // 录制中：显示屏幕变化率、录制间隔、截图数量
                UpdateRecordingStatus();
            }
        }

        /// <summary>
        /// 更新录制状态显示
        /// </summary>
        private void UpdateRecordingStatus()
        {
            try
            {
                if (_isStoppingRecording)
                {
                    if (!string.IsNullOrEmpty(_currentOperation))
                    {
                        UpdateStatusDisplayWithScroll(_currentOperation, System.Windows.Media.Brushes.Orange);
                    }
                    return;
                }

                // 如果截图功能未启用，显示提示信息
                if (!_isScreenshotEnabled)
                {
                    string statusText = $"屏幕变化率: -- | " +
                                       $"录制间隔: {_config.ScreenshotInterval}秒 | " +
                                       $"截图数量: {_screenshotCount} (请启用截图功能)";
                    UpdateStatusDisplayWithScroll(statusText, System.Windows.Media.Brushes.Orange);
                    return;
                }
                
                // 计算距离下次检查的剩余时间
                DateTime now = DateTime.Now;
                TimeSpan elapsed = now - _lastScreenshotTime;
                int remainingSeconds = Math.Max(0, _config.ScreenshotInterval - (int)elapsed.TotalSeconds);
                
                // 始终显示实时的屏幕变化率（比较当前画面和上一次检查时的画面）
                string mainStatusText = $"屏幕变化率: {_currentScreenChangeRate:F2}% | " +
                                       $"录制间隔: {_config.ScreenshotInterval}秒 | " +
                                       $"截图数量: {_screenshotCount}";
                
                // 使用Inlines来设置不同颜色的文本
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Inlines.Clear();
                    StatusBarInfo.Inlines.Add(new System.Windows.Documents.Run(mainStatusText) 
                    { 
                        Foreground = System.Windows.Media.Brushes.Green 
                    });
                    
                    if (remainingSeconds > 0 && _lastScreenshot != null)
                    {
                        // 如果间隔时间未到达，在变化率后面显示剩余时间（使用蓝色以示区分）
                        StatusBarInfo.Inlines.Add(new System.Windows.Documents.Run($" (下次检查: {remainingSeconds}秒后)") 
                        { 
                            Foreground = System.Windows.Media.Brushes.Blue 
                        });
                    }
                    
                    // 检查是否需要滚动
                    CheckAndStartScrolling();
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新录制状态失败", ex);
            }
        }

        /// <summary>
        /// 更新状态栏显示（合并配置信息和状态信息为一行，用不同颜色显示）
        /// </summary>
        /// <param name="statusText">状态文本，如果为null则使用当前状态</param>
        private void UpdateStatusDisplay(string? statusText = null)
        {
            try
            {
                if (StatusBarInfo != null)
                {
                    if (_isStoppingRecording && !string.IsNullOrEmpty(_currentOperation))
                    {
                        UpdateStatusDisplayWithScroll(_currentOperation, System.Windows.Media.Brushes.Orange);
                        return;
                    }

                    // 如果正在录制，显示录制状态
                    if (_isVideoRecording && statusText == null)
                    {
                        UpdateRecordingStatus();
                        return;
                    }

                    // 如果正在停止录制，显示当前操作
                    if (!string.IsNullOrEmpty(_currentOperation))
                    {
                        UpdateStatusDisplayWithScroll(_currentOperation, System.Windows.Media.Brushes.Orange);
                        return;
                    }

                    // 配置信息（蓝色）
                    string regionInfo = HasValidCustomRegion()
                        ? $"{_config.RegionWidth}x{_config.RegionHeight}"
                        : "全屏";
                    string configText = $"{_config.VideoResolutionScale}% {_config.VideoFrameRate}fps | " +
                                       $"{_config.AudioSampleRate / 1000}kHz {_config.AudioBitrate}kbps | {regionInfo}";

                    // 状态信息（绿色表示正常，橙色表示警告，红色表示错误）
                    string status = statusText ?? "就绪";
                    System.Windows.Media.Brush statusColor = System.Windows.Media.Brushes.Green;
                    
                    // 根据状态文本判断颜色
                    if (status.Contains("失败") || status.Contains("错误"))
                    {
                        statusColor = System.Windows.Media.Brushes.Red;
                    }
                    else if (status.Contains("警告") || status.Contains("正在"))
                    {
                        statusColor = System.Windows.Media.Brushes.Orange;
                    }

                    // 使用 Inlines 显示不同颜色的文本
                    StatusBarInfo.Inlines.Clear();
                    StatusBarInfo.Inlines.Add(new System.Windows.Documents.Run(configText) 
                    { 
                        Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 102, 204)) // #0066CC
                    });
                    StatusBarInfo.Inlines.Add(new System.Windows.Documents.Run(" | ") 
                    { 
                        Foreground = System.Windows.Media.Brushes.Gray 
                    });
                    StatusBarInfo.Inlines.Add(new System.Windows.Documents.Run(status) 
                    { 
                        Foreground = statusColor 
                    });
                    
                    // 检查是否需要滚动
                    CheckAndStartScrolling();
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新状态显示失败", ex);
            }
        }

        /// <summary>
        /// 更新状态显示（带滚动支持）
        /// </summary>
        private void UpdateStatusDisplayWithScroll(string statusText, System.Windows.Media.Brush color)
        {
            try
            {
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Inlines.Clear();
                    StatusBarInfo.Inlines.Add(new System.Windows.Documents.Run(statusText) 
                    { 
                        Foreground = color 
                    });
                    
                    // 检查是否需要滚动
                    CheckAndStartScrolling();
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新状态显示失败", ex);
            }
        }

        /// <summary>
        /// 检查并启动滚动动画（跑马灯效果）
        /// </summary>
        private void CheckAndStartScrolling()
        {
            try
            {
                if (StatusBarInfo != null && StatusBarScrollViewer != null)
                {
                    // 使用Dispatcher延迟检查，确保布局已完成
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            // 等待布局完成
                            StatusBarInfo.Measure(new System.Windows.Size(double.PositiveInfinity, double.PositiveInfinity));
                            double textWidth = StatusBarInfo.DesiredSize.Width;
                            double scrollViewerWidth = StatusBarScrollViewer.ActualWidth;
                            
                            // 如果文本宽度超过ScrollViewer宽度，启动滚动动画
                            if (textWidth > scrollViewerWidth && scrollViewerWidth > 0)
                            {
                                StartScrollingAnimation(textWidth - scrollViewerWidth);
                            }
                            else
                            {
                                StopScrollingAnimation();
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteError($"延迟检查滚动状态失败", ex);
                        }
                    }), System.Windows.Threading.DispatcherPriority.Loaded);
                }
            }
            catch (Exception ex)
            {
                WriteError($"检查滚动状态失败", ex);
            }
        }

        private System.Windows.Media.Animation.Storyboard? _scrollingStoryboard;

        /// <summary>
        /// 启动滚动动画（使用Transform实现）
        /// </summary>
        private void StartScrollingAnimation(double scrollDistance)
        {
            try
            {
                // 如果动画已经在运行，不重复启动
                if (_scrollingStoryboard != null && _scrollingStoryboard.GetCurrentState() == System.Windows.Media.Animation.ClockState.Active)
                {
                    return;
                }

                StopScrollingAnimation();

                if (StatusBarInfo == null) return;

                // 使用Transform实现滚动效果
                var transform = new System.Windows.Media.TranslateTransform();
                StatusBarInfo.RenderTransform = transform;

                // 创建滚动动画
                _scrollingStoryboard = new System.Windows.Media.Animation.Storyboard();
                
                var scrollAnimation = new System.Windows.Media.Animation.DoubleAnimation
                {
                    From = 0,
                    To = -(scrollDistance + 50), // 向左滚动
                    Duration = TimeSpan.FromSeconds(5), // 5秒完成一次滚动
                    RepeatBehavior = System.Windows.Media.Animation.RepeatBehavior.Forever,
                    AutoReverse = true
                };

                System.Windows.Media.Animation.Storyboard.SetTarget(scrollAnimation, transform);
                System.Windows.Media.Animation.Storyboard.SetTargetProperty(scrollAnimation, new System.Windows.PropertyPath(System.Windows.Media.TranslateTransform.XProperty));
                
                _scrollingStoryboard.Children.Add(scrollAnimation);
                _scrollingStoryboard.Begin();
            }
            catch (Exception ex)
            {
                WriteError($"启动滚动动画失败", ex);
            }
        }

        /// <summary>
        /// 停止滚动动画
        /// </summary>
        private void StopScrollingAnimation()
        {
            try
            {
                if (_scrollingStoryboard != null)
                {
                    _scrollingStoryboard.Stop();
                    _scrollingStoryboard = null;
                }
                if (StatusBarInfo != null && StatusBarInfo.RenderTransform is System.Windows.Media.TranslateTransform transform)
                {
                    transform.X = 0;
                }
                if (StatusBarScrollViewer != null)
                {
                    StatusBarScrollViewer.ScrollToHorizontalOffset(0);
                }
            }
            catch (Exception ex)
            {
                WriteError($"停止滚动动画失败", ex);
            }
        }

        private bool HasValidCustomRegion()
        {
            return _config.UseCustomRegion &&
                   _config.RegionWidth > 0 &&
                   _config.RegionHeight > 0;
        }

        /// <summary>
        /// 初始化按钮状态：只有选择按钮可用，开始和停止按钮禁用
        /// </summary>
        private void InitializeButtonStates()
        {
            try
            {
                if (BtnStartVideo != null)
                {
                    BtnStartVideo.IsEnabled = false;
                }
                if (BtnStopVideo != null)
                {
                    BtnStopVideo.IsEnabled = false;
                }
                if (BtnSelectRegion != null)
                {
                    BtnSelectRegion.IsEnabled = true;
                }
            }
            catch (Exception ex)
            {
                WriteError($"初始化按钮状态失败", ex);
            }
        }

        /// <summary>
        /// 更新开始按钮的启用状态
        /// 如果没有有效区域，则禁用开始按钮
        /// </summary>
        private void UpdateStartButtonState()
        {
            try
            {
                if (BtnStartVideo != null)
                {
                    // 如果没有有效区域，则禁用开始按钮
                    if (!HasValidCustomRegion())
                    {
                        BtnStartVideo.IsEnabled = false;
                    }
                    else
                    {
                        // 否则启用开始按钮（除非正在录制）
                        BtnStartVideo.IsEnabled = !_isVideoRecording;
                    }
                }
                // 停止按钮始终跟随录制状态
                if (BtnStopVideo != null)
                {
                    BtnStopVideo.IsEnabled = _isVideoRecording;
                }
                // 选择按钮：录制时禁用，否则启用
                if (BtnSelectRegion != null)
                {
                    BtnSelectRegion.IsEnabled = !_isVideoRecording;
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新开始按钮状态失败", ex);
            }
        }

        private void UpdateRegionOverlay()
        {
            try
            {
                // 如果没有有效的自定义区域，则隐藏
                if (!HasValidCustomRegion())
                {
                    _regionHighlightWindow?.HideRegion();
                    return;
                }

                // 有有效区域，显示红色矩形框
                if (_regionHighlightWindow == null)
                {
                    _regionHighlightWindow = new RegionHighlightWindow();
                    // 只在主窗口已显示时设置 Owner
                    if (IsLoaded && IsVisible)
                    {
                        _regionHighlightWindow.Owner = this;
                    }
                }
                else if (IsLoaded && IsVisible && _regionHighlightWindow.Owner == null)
                {
                    // 如果之前没有设置 Owner，现在设置
                    _regionHighlightWindow.Owner = this;
                }

                var rect = new Int32Rect(
                    _config.RegionLeft,
                    _config.RegionTop,
                    _config.RegionWidth,
                    _config.RegionHeight);

                _regionHighlightWindow.ShowRegion(rect);
            }
            catch (Exception ex)
            {
                WriteError($"更新区域高亮失败", ex);
            }
        }

        private void BtnSelectRegion_Click(object sender, RoutedEventArgs e)
        {
            RegionSelectionWindow? selector = null;
            try
            {
                if (_isVideoRecording)
                {
                    UpdateStatusDisplay("请先停止录制，再调整录制区域");
                    return;
                }

                selector = new RegionSelectionWindow
                {
                    Owner = this
                };

                bool? dialogResult = selector.ShowDialog();
                if (dialogResult == true)
                {
                    var rect = selector.SelectedRect;
                    if (rect.Width > 0 && rect.Height > 0)
                    {
                        // 验证区域是否在主屏幕范围内
                        var (screenWidth, screenHeight) = GetPrimaryScreenPixelSize();
                        int regionRight = rect.X + rect.Width;
                        int regionBottom = rect.Y + rect.Height;
                        
                        if (rect.X < 0 || rect.Y < 0 || regionRight > screenWidth || regionBottom > screenHeight)
                        {
                            string warningMsg = $"选择的区域部分超出主屏幕范围！\n" +
                                               $"区域: ({rect.X}, {rect.Y}) 到 ({regionRight}, {regionBottom})\n" +
                                               $"主屏幕: (0, 0) 到 ({screenWidth}, {screenHeight})\n" +
                                               $"录制时只会录制主屏幕范围内的部分。\n\n" +
                                               $"是否继续使用此区域？";
                            
                            WriteWarning(warningMsg);
                            
                            var result = MessageBox.Show(warningMsg, "区域超出范围", 
                                MessageBoxButton.YesNo, MessageBoxImage.Warning);
                            
                            if (result == MessageBoxResult.No)
                            {
                                UpdateStatusDisplay("已取消选择区域");
                                return;
                            }
                        }
                        
                        _config.UseCustomRegion = true;
                        _config.RegionLeft = rect.X;
                        _config.RegionTop = rect.Y;
                        _config.RegionWidth = rect.Width;
                        _config.RegionHeight = rect.Height;
                        _config.Save(_configPath);

                        // 用户选择新区域后，总是显示红色矩形框（不受配置影响）
                        UpdateRegionOverlay();
                        UpdateConfigDisplay();
                        
                        // 选择区域后，更新开始按钮状态
                        UpdateStartButtonState();

                        WriteLine($"选择录制区域: 左上=({rect.X},{rect.Y}), 大小={rect.Width}x{rect.Height}");
                        UpdateStatusDisplay($"已选择区域：{rect.Width}x{rect.Height} @ ({rect.X},{rect.Y})");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"选择录制区域失败", ex);
                UpdateStatusDisplay($"选择录制区域失败：{ex.Message}");
            }
            finally
            {
                // 确保窗口正确关闭和清理
                try
                {
                    if (selector != null && selector.IsVisible)
                    {
                        selector.Close();
                    }
                }
                catch
                {
                    // 忽略关闭错误
                }
            }
        }

        /// <summary>
        /// 恢复最近一次框选的区域
        /// </summary>
        private void BtnRestoreRegion_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isVideoRecording)
                {
                    UpdateStatusDisplay("请先停止录制，再恢复录制区域");
                    return;
                }

                // 检查配置中是否有有效的区域
                if (!HasValidCustomRegion())
                {
                    UpdateStatusDisplay("没有可恢复的区域，请先选择区域");
                    WriteLine("恢复区域失败：配置中没有有效的区域信息");
                    return;
                }

                // 验证区域是否在主屏幕范围内
                var (screenWidth, screenHeight) = GetPrimaryScreenPixelSize();
                int regionRight = _config.RegionLeft + _config.RegionWidth;
                int regionBottom = _config.RegionTop + _config.RegionHeight;

                if (_config.RegionLeft < 0 || _config.RegionTop < 0 ||
                    regionRight > screenWidth || regionBottom > screenHeight)
                {
                    string warningMsg = $"保存的区域部分超出主屏幕范围！\n" +
                                       $"区域: ({_config.RegionLeft}, {_config.RegionTop}) 到 ({regionRight}, {regionBottom})\n" +
                                       $"主屏幕: (0, 0) 到 ({screenWidth}, {screenHeight})\n" +
                                       $"是否继续使用此区域？";

                    WriteWarning(warningMsg);

                    var result = MessageBox.Show(warningMsg, "区域超出范围",
                        MessageBoxButton.YesNo, MessageBoxImage.Warning);

                    if (result == MessageBoxResult.No)
                    {
                        UpdateStatusDisplay("已取消恢复区域");
                        return;
                    }
                }

                // 恢复区域：确保 UseCustomRegion 为 true
                _config.UseCustomRegion = true;
                _config.Save(_configPath);

                // 更新界面显示
                UpdateRegionOverlay();
                UpdateConfigDisplay();
                UpdateStartButtonState();

                WriteLine($"恢复录制区域: 左上=({_config.RegionLeft},{_config.RegionTop}), 大小={_config.RegionWidth}x{_config.RegionHeight}");
                UpdateStatusDisplay($"已恢复区域：{_config.RegionWidth}x{_config.RegionHeight} @ ({_config.RegionLeft},{_config.RegionTop})");
            }
            catch (Exception ex)
            {
                WriteError($"恢复录制区域失败", ex);
                UpdateStatusDisplay($"恢复录制区域失败：{ex.Message}");
            }
        }

        private (int width, int height) GetPrimaryScreenPixelSize()
        {
            int width = GetSystemMetrics(SM_CXSCREEN);
            int height = GetSystemMetrics(SM_CYSCREEN);

            if (width <= 0 || height <= 0)
            {
                // 回退到 WPF 的 SystemParameters（单位为 DIP，需要乘 DPI 比例，但通常足够）
                width = (int)SystemParameters.PrimaryScreenWidth;
                height = (int)SystemParameters.PrimaryScreenHeight;
            }

            return (width, height);
        }

        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        /// <summary>
        /// 初始化界面上的设置输入框
        /// </summary>
        private void InitializeSettingsControls()
        {
            try
            {
                if (TxtScreenChangeRate != null)
                {
                    TxtScreenChangeRate.Text = _config.ScreenChangeRate.ToString("F2");
                }
                if (TxtScreenshotInterval != null)
                {
                    TxtScreenshotInterval.Text = _config.ScreenshotInterval.ToString();
                }
            }
            catch (Exception ex)
            {
                WriteError($"初始化设置控件失败", ex);
            }
        }

        /// <summary>
        /// 屏幕变化率输入框文本更改事件
        /// </summary>
        private void TxtScreenChangeRate_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            try
            {
                if (TxtScreenChangeRate != null && double.TryParse(TxtScreenChangeRate.Text, out double value))
                {
                    // 验证范围：1-1000
                    if (value < 1)
                    {
                        value = 1;
                        TxtScreenChangeRate.Text = "1";
                    }
                    else if (value > 1000)
                    {
                        value = 1000;
                        TxtScreenChangeRate.Text = "1000";
                    }
                    
                    // 更新配置并保存
                    if (Math.Abs(_config.ScreenChangeRate - value) > 0.01) // 避免频繁保存（浮点数比较）
                    {
                        _config.ScreenChangeRate = value;
                        _config.Save(_configPath);
                        WriteLine($"屏幕变化率已更新: {value:F2}%");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新屏幕变化率失败", ex);
            }
        }

        /// <summary>
        /// 截图间隔输入框文本更改事件
        /// </summary>
        private void TxtScreenshotInterval_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            try
            {
                if (TxtScreenshotInterval != null && int.TryParse(TxtScreenshotInterval.Text, out int value))
                {
                    // 验证范围：1-65535
                    if (value < 1)
                    {
                        value = 1;
                        TxtScreenshotInterval.Text = "1";
                    }
                    else if (value > 65535)
                    {
                        value = 65535;
                        TxtScreenshotInterval.Text = "65535";
                    }
                    
                    // 更新配置并保存
                    if (_config.ScreenshotInterval != value)
                    {
                        _config.ScreenshotInterval = value;
                        _config.Save(_configPath);
                        WriteLine($"截图间隔已更新: {value}秒");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新截图间隔失败", ex);
            }
        }

        /// <summary>
        /// 更新日志菜单项状态
        /// </summary>
        private void UpdateLogMenuItems()
        {
            try
            {
                // 更新日志开关状态
                if (MenuLogEnabled != null)
                {
                    MenuLogEnabled.IsChecked = _config.LogEnabled == 1;
                }
                if (MenuLogDisabled != null)
                {
                    MenuLogDisabled.IsChecked = _config.LogEnabled == 0;
                }

                // 更新日志文件模式状态
                if (MenuLogFileOverwrite != null)
                {
                    MenuLogFileOverwrite.IsChecked = _config.LogFileMode == 0;
                }
                if (MenuLogFileAppend != null)
                {
                    MenuLogFileAppend.IsChecked = _config.LogFileMode == 1;
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新日志菜单项状态失败", ex);
            }
        }

        /// <summary>
        /// 更新PPT和PDF菜单项状态
        /// </summary>
        private void UpdatePPTAndPDFMenuItems()
        {
            try
            {
                // 如果截图功能未开启，禁用PPT和PDF菜单项
                bool isEnabled = _isScreenshotEnabled;
                
                if (MenuGeneratePPT != null)
                {
                    MenuGeneratePPT.IsChecked = _config.GeneratePPT;
                    MenuGeneratePPT.IsEnabled = isEnabled;
                }
                if (MenuGeneratePDF != null)
                {
                    MenuGeneratePDF.IsChecked = _config.GeneratePDF;
                    MenuGeneratePDF.IsEnabled = isEnabled;
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新PPT和PDF菜单项状态失败", ex);
            }
        }

        /// <summary>
        /// 保存日志配置
        /// </summary>
        private void SaveLogConfig()
        {
            try
            {
                _config.Save(_configPath);
                Logger.SetLogFileMode(_config.LogFileMode);
                // 如果切换到覆盖模式，需要重新设置日志目录以清空文件
                if (_config.LogFileMode == 0)
                {
                    Logger.SetLogDirectory(_workDir, applyMode: true);
                }
                // 只有在日志启用时才写入日志（避免在禁用时写入）
                if (_config.LogEnabled == 1)
                {
                    Logger.WriteLine($"日志配置已更新: 状态={(_config.LogEnabled == 1 ? "启用" : "禁用")}, 模式={(_config.LogFileMode == 0 ? "覆盖" : "叠加")}");
                }
            }
            catch (Exception ex)
            {
                WriteError($"保存日志配置失败", ex);
            }
        }

        private void MenuLogEnabled_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.LogEnabled = 1;
                Logger.Enabled = true;
                UpdateLogMenuItems();
                SaveLogConfig();
                UpdateStatusDisplay("日志已启用");
            }
            catch (Exception ex)
            {
                WriteError($"启用日志失败", ex);
            }
        }

        private void MenuLogDisabled_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.LogEnabled = 0;
                Logger.Enabled = false;
                UpdateLogMenuItems();
                SaveLogConfig();
                UpdateStatusDisplay("日志已禁用");
            }
            catch (Exception ex)
            {
                WriteError($"禁用日志失败", ex);
            }
        }

        private void MenuLogFileOverwrite_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.LogFileMode = 0;
                UpdateLogMenuItems();
                SaveLogConfig();
                UpdateStatusDisplay("日志文件模式：覆盖");
            }
            catch (Exception ex)
            {
                WriteError($"设置日志文件覆盖模式失败", ex);
            }
        }

        private void MenuLogFileAppend_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.LogFileMode = 1;
                UpdateLogMenuItems();
                SaveLogConfig();
                UpdateStatusDisplay("日志文件模式：叠加");
            }
            catch (Exception ex)
            {
                WriteError($"设置日志文件叠加模式失败", ex);
            }
        }

        private void MenuScreenshot_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _isScreenshotEnabled = MenuScreenshot.IsChecked;
                
                if (_isScreenshotEnabled)
                {
                    // 启动截图定时器
                    if (_screenshotTimer != null)
                    {
                        // 设置定时器间隔为1秒，但实际检查逻辑会按照配置的间隔执行
                        _screenshotTimer.Interval = TimeSpan.FromSeconds(1);
                        _screenshotTimer.Start();
                        _lastScreenshotTime = DateTime.Now;
                        // 立即初始化 _lastScreenshot，作为第一次检查的基准画面
                        UpdateLastScreenshot();
                        
                        // 重置PPT/PDF初始化状态，准备创建新的PPT/PDF文件
                        _pptPdfInitialized = false;
                        _pptGenerator?.Dispose();
                        _pdfGenerator?.Dispose();
                        _pptGenerator = null;
                        _pdfGenerator = null;
                        
                        WriteLine("截图功能已启用");
                        UpdateStatusDisplay("截图功能已启用");
                    }
                }
                else
                {
                    // 停止截图定时器
                    if (_screenshotTimer != null)
                    {
                        _screenshotTimer.Stop();
                    }
                    _lastScreenshot?.Dispose();
                    _lastScreenshot = null;
                    
                    // 完成并释放PPT/PDF生成器
                    FinalizePPTAndPDF();
                    
                    WriteLine("截图功能已禁用");
                    UpdateStatusDisplay("截图功能已禁用");
                }
                
                // 更新PPT和PDF菜单项的启用状态
                UpdatePPTAndPDFMenuItems();
            }
            catch (Exception ex)
            {
                WriteError($"切换截图功能失败", ex);
                UpdateStatusDisplay($"切换截图功能失败：{ex.Message}");
            }
        }

        private void MenuGeneratePPT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.GeneratePPT = MenuGeneratePPT.IsChecked;
                _config.Save(_configPath);
                WriteLine($"生成PPT: {(_config.GeneratePPT ? "启用" : "禁用")}");
                UpdateStatusDisplay($"生成PPT: {(_config.GeneratePPT ? "已启用" : "已禁用")}");
            }
            catch (Exception ex)
            {
                WriteError($"切换生成PPT功能失败", ex);
                UpdateStatusDisplay($"切换生成PPT功能失败：{ex.Message}");
            }
        }

        private void MenuGeneratePDF_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _config.GeneratePDF = MenuGeneratePDF.IsChecked;
                _config.Save(_configPath);
                WriteLine($"生成PDF: {(_config.GeneratePDF ? "启用" : "禁用")}");
                UpdateStatusDisplay($"生成PDF: {(_config.GeneratePDF ? "已启用" : "已禁用")}");
            }
            catch (Exception ex)
            {
                WriteError($"切换生成PDF功能失败", ex);
                UpdateStatusDisplay($"切换生成PDF功能失败：{ex.Message}");
            }
        }


        private void ScreenshotTimer_Tick(object? sender, EventArgs e)
        {
            try
            {
                if (!_isScreenshotEnabled) return;

                DateTime now = DateTime.Now;
                TimeSpan elapsed = now - _lastScreenshotTime;
                
                // 持续计算屏幕变化率（用于实时显示）
                // 这比较的是"上一次检查时的画面"（间隔时间前的画面）和"当前画面"的变化率
                // 注意：这里不更新 _lastScreenshot，只计算变化率用于显示
                double changeRate = CalculateScreenChangeRateForDisplay();
                _currentScreenChangeRate = changeRate;
                
                // 检查是否到了截图间隔时间
                if (elapsed.TotalSeconds >= _config.ScreenshotInterval)
                {
                    // 间隔时间到达，检查是否需要截图
                    if (changeRate >= _config.ScreenChangeRate || _lastScreenshot == null)
                    {
                        // 屏幕变化率超过阈值，执行截图
                        CaptureScreenshot();
                    }
                    
                    // 间隔时间到达，更新 _lastScreenshot 为当前画面（保存为"上一次检查时的画面"）
                    UpdateLastScreenshot();
                    _lastScreenshotTime = now;
                }
            }
            catch (Exception ex)
            {
                WriteError($"截图定时器处理失败", ex);
            }
        }

        /// <summary>
        /// 计算屏幕变化率（仅用于显示，不更新 _lastScreenshot）
        /// 比较"上一次检查时的画面"（间隔时间前的画面）和"当前画面"的变化率
        /// </summary>
        private double CalculateScreenChangeRateForDisplay()
        {
            Bitmap? currentScreenshot = null;
            try
            {
                // 获取当前屏幕截图
                currentScreenshot = CaptureScreenBitmap();
                
                if (_lastScreenshot == null)
                {
                    // 第一次，还没有基准画面，返回最大值
                    currentScreenshot?.Dispose();
                    return 1000;
                }

                // 检查两张图片的尺寸是否匹配
                if (currentScreenshot.Width != _lastScreenshot.Width || 
                    currentScreenshot.Height != _lastScreenshot.Height)
                {
                    // 尺寸不匹配（可能是区域选择改变了），返回最大值表示需要更新
                    currentScreenshot?.Dispose();
                    return 1000;
                }

                // 比较两张图片的差异（参考Python代码的pixel_diff函数）
                // Python逻辑：np.any(diff > 20, axis=2) - 如果任何一个颜色通道的差值超过20，认为像素变化
                int changedPixels = 0;

                // 为了提高性能，使用采样方式（每5个像素采样一次）
                // 但计算方式与Python保持一致：检查任何一个通道是否超过20
                int sampleRate = 5;
                int sampledPixels = 0;
                
                // 使用较小的尺寸确保不越界
                int width = Math.Min(currentScreenshot.Width, _lastScreenshot.Width);
                int height = Math.Min(currentScreenshot.Height, _lastScreenshot.Height);
                
                for (int y = 0; y < height; y += sampleRate)
                {
                    for (int x = 0; x < width; x += sampleRate)
                    {
                        // 确保坐标在有效范围内
                        if (x >= currentScreenshot.Width || y >= currentScreenshot.Height ||
                            x >= _lastScreenshot.Width || y >= _lastScreenshot.Height)
                        {
                            continue;
                        }
                        
                        Color currentColor = currentScreenshot.GetPixel(x, y);
                        Color lastColor = _lastScreenshot.GetPixel(x, y);
                        
                        // 计算每个颜色通道的差值（参考Python: np.abs(arr1 - arr2)）
                        int diffR = Math.Abs(currentColor.R - lastColor.R);
                        int diffG = Math.Abs(currentColor.G - lastColor.G);
                        int diffB = Math.Abs(currentColor.B - lastColor.B);
                        
                        // Python逻辑：np.any(diff > 20, axis=2) - 如果任何一个通道差值超过20，认为像素变化
                        if (diffR > 20 || diffG > 20 || diffB > 20)
                        {
                            changedPixels++;
                        }
                        sampledPixels++;
                    }
                }

                // 计算变化率（百分比）：基于采样像素的变化率
                // Python逻辑：rate = (changed_count / total) * 100
                double sampledChangeRate = (changedPixels / (double)sampledPixels) * 100.0;
                
                // 释放当前截图（不更新 _lastScreenshot）
                currentScreenshot?.Dispose();
                
                return sampledChangeRate;
            }
            catch (Exception ex)
            {
                WriteError($"计算屏幕变化率失败", ex);
                currentScreenshot?.Dispose(); // 确保在异常时释放资源
                return 0;
            }
        }

        /// <summary>
        /// 更新 _lastScreenshot 为当前画面（仅在间隔时间到达时调用）
        /// </summary>
        private void UpdateLastScreenshot()
        {
            try
            {
                Bitmap? currentScreenshot = CaptureScreenBitmap();
                _lastScreenshot?.Dispose();
                _lastScreenshot = currentScreenshot;
            }
            catch (Exception ex)
            {
                WriteError($"更新上次截图失败", ex);
            }
        }

        private Bitmap CaptureScreenBitmap()
        {
            try
            {
                // 如果设置了框选区域，只捕获框选区域；否则捕获整个屏幕
                if (HasValidCustomRegion())
                {
                    // 向内收缩几个像素，避开红色框的边框（红色框边框宽度为2像素，收缩3像素更安全）
                    const int borderOffset = 3;
                    int offsetX = _config.RegionLeft + borderOffset;
                    int offsetY = _config.RegionTop + borderOffset;
                    int width = _config.RegionWidth - borderOffset * 2;  // 左右各收缩
                    int height = _config.RegionHeight - borderOffset * 2; // 上下各收缩
                    
                    // 确保宽度和高度为正数
                    if (width <= 0) width = 1;
                    if (height <= 0) height = 1;
                    
                    Bitmap bitmap = new Bitmap(width, height);
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(
                            offsetX, 
                            offsetY, 
                            0, 
                            0, 
                            new System.Drawing.Size(width, height)
                        );
                    }
                    return bitmap;
                }
                else
                {
                    // 没有框选区域，捕获整个屏幕
                    var (screenWidth, screenHeight) = GetPrimaryScreenPixelSize();
                    Bitmap bitmap = new Bitmap(screenWidth, screenHeight);
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(0, 0, 0, 0, new System.Drawing.Size(screenWidth, screenHeight));
                    }
                    return bitmap;
                }
            }
            catch (Exception ex)
            {
                WriteError($"捕获屏幕位图失败", ex);
                throw;
            }
        }

        private void CaptureScreenshot()
        {
            try
            {
                int width, height;
                Bitmap bitmap;
                
                // 如果设置了框选区域，只捕获框选区域；否则捕获整个屏幕
                if (HasValidCustomRegion())
                {
                    // 向内收缩几个像素，避开红色框的边框（红色框边框宽度为2像素，收缩3像素更安全）
                    const int borderOffset = 3;
                    int offsetX = _config.RegionLeft + borderOffset;
                    int offsetY = _config.RegionTop + borderOffset;
                    width = _config.RegionWidth - borderOffset * 2;  // 左右各收缩
                    height = _config.RegionHeight - borderOffset * 2; // 上下各收缩
                    
                    // 确保宽度和高度为正数
                    if (width <= 0) width = 1;
                    if (height <= 0) height = 1;
                    
                    bitmap = new Bitmap(width, height);
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(
                            offsetX, 
                            offsetY, 
                            0, 
                            0, 
                            new System.Drawing.Size(width, height)
                        );
                    }
                }
                else
                {
                    var (screenWidth, screenHeight) = GetPrimaryScreenPixelSize();
                    width = screenWidth;
                    height = screenHeight;
                    bitmap = new Bitmap(width, height);
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen(0, 0, 0, 0, new System.Drawing.Size(width, height));
                    }
                }
                
                using (bitmap)
                {
                    // 生成文件名：Picyymmddhhmmss.jpg
                    string fileName = $"Pic{DateTime.Now:yyMMddHHmmss}.jpg";
                    string filePath = Path.Combine(_screenshotDir, fileName);
                    
                    // 保存为JPG格式
                    ImageCodecInfo? jpegCodec = ImageCodecInfo.GetImageEncoders()
                        .FirstOrDefault(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
                    
                    if (jpegCodec != null)
                    {
                        EncoderParameters encoderParams = new EncoderParameters(1);
                        encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, 90L); // 90%质量
                        bitmap.Save(filePath, jpegCodec, encoderParams);
                        encoderParams.Dispose();
                    }
                    else
                    {
                        // 如果没有找到JPEG编码器，使用默认方式保存
                        bitmap.Save(filePath, ImageFormat.Jpeg);
                    }
                    
                    WriteLine($"截图已保存: {filePath}");
                    
                    // 增加截图计数
                    _screenshotCount++;
                    
                    // 如果启用了PPT或PDF生成，立即添加到PPT/PDF
                    if (_config.GeneratePPT || _config.GeneratePDF)
                    {
                        InitializePPTAndPDF(width, height);
                        
                        if (_config.GeneratePPT && _pptGenerator != null)
                        {
                            try
                            {
                                _pptGenerator.AddImage(filePath);
                            }
                            catch (Exception ex)
                            {
                                WriteError($"添加图片到PPT失败", ex);
                            }
                        }
                        
                        if (_config.GeneratePDF && _pdfGenerator != null)
                        {
                            try
                            {
                                _pdfGenerator.AddImage(filePath);
                            }
                            catch (Exception ex)
                            {
                                WriteError($"添加图片到PDF失败", ex);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"截图保存失败", ex);
            }
        }

        private void InitializePPTAndPDF(int width, int height)
        {
            if (_pptPdfInitialized) return;
            
            try
            {
                if (_config.GeneratePPT)
                {
                    string timestamp = DateTime.Now.ToString("yyMMddHHmmss");
                    _pptFilePath = Path.Combine(_screenshotDir, $"PPT{timestamp}.pptx");
                    _pptGenerator = new PPTGenerator(_pptFilePath, width, height);
                    _pptGenerator.Initialize();
                    WriteLine($"PPT生成器已初始化: {_pptFilePath}");
                }
                
                if (_config.GeneratePDF)
                {
                    string timestamp = DateTime.Now.ToString("yyMMddHHmmss");
                    _pdfFilePath = Path.Combine(_screenshotDir, $"PDF{timestamp}.pdf");
                    _pdfGenerator = new PDFGenerator(_pdfFilePath, width, height);
                    _pdfGenerator.Initialize();
                    WriteLine($"PDF生成器已初始化: {_pdfFilePath}");
                }
                
                _pptPdfInitialized = true;
            }
            catch (Exception ex)
            {
                WriteError($"初始化PPT/PDF生成器失败", ex);
            }
        }

        private void FinalizePPTAndPDF()
        {
            try
            {
                if (_config.GeneratePPT && _pptGenerator != null)
                {
                    _pptGenerator.Finish();
                    WriteLine($"PPT文件已生成: {_pptFilePath}");
                }
                
                if (_config.GeneratePDF && _pdfGenerator != null)
                {
                    _pdfGenerator.Finish();
                    WriteLine($"PDF文件已生成: {_pdfFilePath}");
                }
            }
            catch (Exception ex)
            {
                WriteError($"完成PPT/PDF生成失败", ex);
            }
            finally
            {
                _pptGenerator?.Dispose();
                _pdfGenerator?.Dispose();
                _pptGenerator = null;
                _pdfGenerator = null;
                _pptPdfInitialized = false;
            }
        }
    }
}
