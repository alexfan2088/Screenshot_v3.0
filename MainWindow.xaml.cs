using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Drawing;
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

        public MainWindow()
        {
            InitializeComponent();
            
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

            // 加载配置
            _config = RecordingConfig.Load(_configPath);
            
            // 初始化日志系统
            Logger.Enabled = _config.LogEnabled == 1;
            // 设置日志文件目录为工作目录（与视频、音频文件同一目录）
            // SetLogDirectory 会自动清空日志文件（覆盖模式）
            Logger.SetLogDirectory(_workDir);
            Logger.WriteLine("程序启动");
            Logger.WriteLine($"日志状态: {(Logger.Enabled ? "启用" : "禁用")}");
            
            UpdateConfigDisplay();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            // 立即允许关闭，不阻塞
            e.Cancel = false;

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
                    if (StatusBarInfo != null)
                    {
                        StatusBarInfo.Text = "设置已保存";
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"打开设置失败：{ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMsg += $"\n详细：{ex.InnerException.Message}";
                }
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = errorMsg;
                }
                WriteError($"打开设置窗口异常", ex);
            }
        }

        private void BtnStartVideo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isVideoRecording) return;

                var videoPath = Path.Combine(_workDir, $"video{DateTime.Now:yyMMddHHmmss}.mp4");
                
                // 计算分辨率
                double resolutionScale = _config.VideoResolutionScale / 100.0;
                int screenWidth = (int)SystemParameters.PrimaryScreenWidth;
                int screenHeight = (int)SystemParameters.PrimaryScreenHeight;
                int videoWidth = (int)(screenWidth * resolutionScale);
                int videoHeight = (int)(screenHeight * resolutionScale);

                // 创建视频编码器（不再需要 VideoRecorder，FFmpeg 直接录制）
                _videoEncoder = new VideoEncoder(videoPath, _config);

                // 初始化编码器（offsetX 和 offsetY 都是 0，从屏幕左上角开始）
                _videoEncoder.Initialize(videoWidth, videoHeight, _config.VideoFrameRate, 
                    _config.AudioSampleRate, 2, offsetX: 0, offsetY: 0);

                // 开始音频录制（使用 NAudio WasapiLoopbackCapture，无需虚拟声卡）
                // 注意：无论是否有声音输出，都要录制音频（包括静音）
                // AudioRecorder.Start() 会创建 temp_{filename}.wav 文件
                var timestamp = DateTime.Now.ToString("yyMMddHHmmss");
                var audioOutputPath = Path.Combine(_workDir, $"audio_{timestamp}.m4a"); // 最终输出路径（不会被使用）
                var expectedAudioPath = Path.Combine(_workDir, $"temp_audio_{timestamp}.wav"); // AudioRecorder 实际创建的文件
                WriteLine($"========== 开始录制 ==========");
                WriteLine($"视频文件: {videoPath}");
                WriteLine($"开始音频录制（NAudio WasapiLoopbackCapture，即使静音也会录制），输出路径: {audioOutputPath}");
                WriteLine($"预期音频文件: {expectedAudioPath}");
                _audioRecorder.Start(audioOutputPath);
                
                // 等待音频录制启动
                System.Threading.Thread.Sleep(300);
                
                // 启动 FFmpeg 录制视频
                // VideoEncoder.Start() 会尝试直接录制有声视频，如果失败则只录制视频
                WriteLine($"启动 FFmpeg 录制视频...");
                _videoEncoder.Start();

                _isVideoRecording = true;
                BtnStartVideo.IsEnabled = false;
                BtnStopVideo.IsEnabled = true;

                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = $"开始录制视频 → {videoPath} | 分辨率: {videoWidth}x{videoHeight} @ {_config.VideoFrameRate}fps";
                }
            }
            catch (Exception ex)
            {
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = $"开始录制视频失败：{ex.Message}";
                }
            }
        }

        private void BtnStopVideo_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!_isVideoRecording) return;

                // 先禁用按钮，防止重复点击
                BtnStopVideo.IsEnabled = false;
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = "正在停止录制...";
                }

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
                        
                        // 立即停止音频录制（与视频同时停止）
                        WriteLine($"停止音频录制");
                        _audioRecorder.Stop();
                        
                        // 等待音频文件写入完成
                        System.Threading.Thread.Sleep(1000);
                        
                        // 设置音频文件路径（必须在调用 Finish() 之前设置）
                        // AudioRecorder 创建的文件名是 temp_{filename}.wav
                        // 查找所有 temp_*.wav 文件
                        var tempAudioFiles = Directory.GetFiles(_workDir, "temp_*.wav");
                        WriteLine($"找到 {tempAudioFiles.Length} 个临时音频文件（temp_*.wav）");
                        
                        if (tempAudioFiles.Length == 0)
                        {
                            // 也尝试查找 temp_audio_*.wav（以防万一）
                            tempAudioFiles = Directory.GetFiles(_workDir, "temp_audio_*.wav");
                            WriteLine($"找到 {tempAudioFiles.Length} 个临时音频文件（temp_audio_*.wav）");
                        }
                        
                        if (tempAudioFiles.Length > 0)
                        {
                            var latestAudioFile = tempAudioFiles.OrderByDescending(f => File.GetCreationTime(f)).First();
                            var audioFileInfo = new FileInfo(latestAudioFile);
                            WriteLine($"使用音频文件: {latestAudioFile}, 大小: {audioFileInfo.Length} 字节");
                            
                            if (audioFileInfo.Length > 0)
                            {
                                WriteLine($"准备设置音频文件到 VideoEncoder: {latestAudioFile}");
                                _videoEncoder?.SetAudioFile(latestAudioFile);
                                WriteLine($"音频文件已设置到 VideoEncoder");
                            }
                            else
                            {
                                WriteWarning($"音频文件大小为 0，跳过合并");
                            }
                        }
                        else
                        {
                            WriteWarning($"未找到临时音频文件，列出所有文件:");
                            try
                            {
                                var allFiles = Directory.GetFiles(_workDir);
                                foreach (var file in allFiles)
                                {
                                    WriteLine($"  - {file}");
                                }
                            }
                            catch { }
                        }
                        
                        // 现在停止视频录制并完成编码（会合并音频）
                        // 注意：必须在设置音频文件路径之后调用 Finish()
                        WriteLine($"准备停止视频录制（FFmpeg）并完成编码");
                        _videoEncoder?.Finish(); // 这会发送 'q' 给 FFmpeg，等待退出，然后合并音频
                        WriteLine($"视频编码完成");

                // 释放资源
                _videoEncoder?.Dispose();

                        // 在 UI 线程更新界面
                        Dispatcher.Invoke(() =>
                        {
                            _videoEncoder = null;
                            _isVideoRecording = false;
                            BtnStartVideo.IsEnabled = true;
                            
                            if (StatusBarInfo != null)
                            {
                                StatusBarInfo.Text = "视频录制已停止（编码完成需几秒）";
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        WriteError($"停止录制失败", ex);
                        Dispatcher.Invoke(() =>
                        {
                            if (StatusBarInfo != null)
                            {
                                StatusBarInfo.Text = $"停止录制失败：{ex.Message}";
                            }
                            BtnStartVideo.IsEnabled = true;
                            BtnStopVideo.IsEnabled = true;
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                if (StatusBarInfo != null)
                {
                    StatusBarInfo.Text = $"停止录制视频失败：{ex.Message}";
                }
                BtnStartVideo.IsEnabled = true;
                BtnStopVideo.IsEnabled = true;
            }
        }

        private void OnAudioSampleAvailable(byte[]? audioData, int bytesRecorded)
        {
            // 音频数据由 AudioRecorder 直接保存到文件，FFmpeg 会读取该文件
            // 不再需要通过事件传递音频数据
            // 这个方法保留是为了兼容 AudioRecorder 的事件，但不需要做任何处理
        }


        private void UpdateConfigDisplay()
        {
            try
            {
                if (StatusBarConfig != null)
                {
                    StatusBarConfig.Text = $"视频：{_config.VideoResolutionScale}%分辨率，{_config.VideoFrameRate}fps，H.264 | " +
                                         $"音频：{_config.AudioSampleRate / 1000}kHz，{_config.AudioBitrate}kbps";
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新配置显示失败", ex);
            }
        }
    }
}
