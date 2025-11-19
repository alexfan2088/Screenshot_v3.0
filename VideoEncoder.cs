using System;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 视频编码器（使用 FFmpeg gdigrab 直接录制屏幕）
    /// 参考 Python 实现，使用 gdigrab 直接捕获屏幕，避免花屏问题
    /// </summary>
    public sealed class VideoEncoder : IDisposable
    {
        private readonly string _outputPath;
        private readonly RecordingConfig _config;
        private bool _isInitialized;
        private Process? _ffmpegProcess;
        private int _videoWidth;
        private int _videoHeight;
        private int _frameRate;
        private int _audioSampleRate;
        private readonly object _lockObject = new object();
        private string? _ffmpegPath;
        private string? _tempAudioPath;
        private int _offsetX;
        private int _offsetY;
        private bool _hasAudioInVideo; // 标记视频中是否已包含音频（FFmpeg 直接录制）
        private bool _hasRequestedStop; // 标记是否已发送停止信号

        public VideoEncoder(string outputPath, RecordingConfig config)
        {
            _outputPath = outputPath;
            _config = config;
        }

        /// <summary>
        /// 初始化编码器
        /// </summary>
        /// <param name="videoWidth">视频宽度</param>
        /// <param name="videoHeight">视频高度</param>
        /// <param name="frameRate">帧率</param>
        /// <param name="audioSampleRate">音频采样率</param>
        /// <param name="audioChannels">音频通道数</param>
        /// <param name="offsetX">屏幕偏移 X（左上角 X 坐标）</param>
        /// <param name="offsetY">屏幕偏移 Y（左上角 Y 坐标）</param>
        public void Initialize(int videoWidth, int videoHeight, int frameRate, int audioSampleRate, int audioChannels, int offsetX = 0, int offsetY = 0)
        {
            if (_isInitialized) return;

            try
            {
                _videoWidth = videoWidth;
                _videoHeight = videoHeight;
                _frameRate = frameRate;
                _audioSampleRate = audioSampleRate;
                _offsetX = offsetX;
                _offsetY = offsetY;

                // 确保宽度和高度是 2 的倍数（H.264 要求）
                _videoWidth = _videoWidth + (_videoWidth % 2);
                _videoHeight = _videoHeight + (_videoHeight % 2);

                // 查找 FFmpeg
                _ffmpegPath = FindFfmpeg();
                if (string.IsNullOrEmpty(_ffmpegPath))
                {
                    throw new FileNotFoundException("未找到 FFmpeg。请将 ffmpeg.exe 放在程序目录中，或添加到系统 PATH。");
                }

                // 不在这里设置 _tempAudioPath，等待 SetAudioFile 设置
                   _tempAudioPath = null;
                   _hasAudioInVideo = false; // 初始化时假设没有音频
                   _hasRequestedStop = false; // 初始化时未发送停止信号

                _isInitialized = true;

                WriteLine($"视频编码器初始化: {_videoWidth}x{_videoHeight} @ {_frameRate}fps, 偏移: ({_offsetX}, {_offsetY})");
            }
            catch (Exception ex)
            {
                WriteError($"初始化视频编码器失败", ex);
                throw;
            }
        }

        /// <summary>
        /// 开始录制（使用 FFmpeg gdigrab 直接录制屏幕）
        /// </summary>
        /// <param name="audioFilePath">音频文件路径（如果存在）</param>
        public void Start(string? audioFilePath = null)
        {
            if (!_isInitialized) return;

            try
            {
                _tempAudioPath = audioFilePath;

                // 计算视频码率（Mbps）
                int videoBitrateMbps = _config.GetVideoBitrateMbps();
                int videoBitrateKbps = videoBitrateMbps * 1000;

                // 方案 A（推荐）：尝试直接录制有声视频
                // 如果找不到音频设备，回退到只录制视频（需要后续合并）
                string audioDevice = GetSystemAudioDevice();
                string arguments;

                if (!string.IsNullOrEmpty(audioDevice))
                {
                    // ✅ 同时录屏幕 + 系统音频（强烈推荐）
                    arguments = BuildFfmpegCommandWithAudio(videoBitrateKbps, audioDevice);
                    _hasAudioInVideo = true;
                    WriteLine($"✓ 录制方式: FFmpeg 直接生成 MP4（视频+音频同步录制）");
                    WriteLine($"  音频设备: {audioDevice}");
                    WriteLine($"  说明: 视频和音频在同一 FFmpeg 进程中录制，时间戳同步，无需后续合并");
                }
                else
                {
                    // 找不到音频设备，退回到只录视频（需要后续合并）
                    arguments = BuildFfmpegCommandWithoutAudio(videoBitrateKbps);
                    _hasAudioInVideo = false;
                    WriteLine($"✓ 录制方式: 分步录制后合并");
                    WriteLine($"  视频: FFmpeg 录制（gdigrab）");
                    WriteLine($"  音频: NAudio WasapiLoopbackCapture 录制（即使静音也会录制）");
                    WriteLine($"  说明: 录制完成后使用 FFmpeg 合并音频到 MP4");
                }

                var processInfo = new ProcessStartInfo
                {
                    FileName = _ffmpegPath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardInput = true,  // 需要标准输入来发送 'q' 命令停止录制
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                _ffmpegProcess = Process.Start(processInfo);
                if (_ffmpegProcess == null)
                {
                    throw new Exception("无法启动 FFmpeg 进程");
                }

                // 启动错误输出读取线程（避免缓冲区满导致进程阻塞）
                Task.Run(() =>
                {
                    try
                    {
                        string? errorOutput = _ffmpegProcess.StandardError.ReadToEnd();
                        if (!string.IsNullOrEmpty(errorOutput))
                        {
                            WriteLine($"FFmpeg 输出: {errorOutput}");
                        }
                    }
                    catch { }
                });

                WriteLine($"FFmpeg 录制已启动: {_outputPath}");
            }
            catch (Exception ex)
            {
                WriteError($"启动录制失败", ex);
                throw;
            }
        }

        /// <summary>
        /// 构建带音频的 FFmpeg 命令（使用 dshow 直接录制系统音频）
        /// </summary>
        private string BuildFfmpegCommandWithAudio(int videoBitrateKbps, string audioDevice)
        {
            // 参考 Python 实现：使用 gdigrab 录制屏幕，dshow 录制系统音频
            
            return $"-hide_banner -nostats -loglevel warning " +
                   $"-thread_queue_size 1024 " +
                   $"-f gdigrab " +
                   $"-framerate {_frameRate} " +
                   $"-offset_x {_offsetX} " +
                   $"-offset_y {_offsetY} " +
                   $"-video_size {_videoWidth}x{_videoHeight} " +
                   $"-use_wallclock_as_timestamps 1 " +
                   $"-i desktop " +
                   $"-thread_queue_size 1024 " +
                   $"-f dshow " +
                   $"-rtbufsize 256M " +
                   $"-use_wallclock_as_timestamps 1 " +
                   $"-i \"audio={audioDevice}\" " +
                   $"-c:v libx264 " +
                   $"-preset medium " +
                   $"-b:v {videoBitrateKbps}k " +
                   $"-pix_fmt yuv420p " +
                   $"-c:a aac " +
                   $"-b:a {_config.AudioBitrate}k " +
                   $"-movflags +faststart " +
                   $"-r {_frameRate} " +
                   $"-g {_frameRate * 2} " +
                   $"-keyint_min {_frameRate} " +
                   $"-sc_threshold 0 " +
                   $"-threads 2 " +
                   $"-maxrate {videoBitrateKbps * 2}k " +
                   $"-bufsize {videoBitrateKbps * 4}k " +
                   $"-avoid_negative_ts make_zero " +
                   $"-fflags +genpts " +
                   $"-shortest " +
                   $"-y \"{_outputPath}\"";
        }
        
        /// <summary>
        /// 获取系统音频设备名称
        /// </summary>
        private string GetSystemAudioDevice()
        {
            // 常见的系统音频设备名称（按优先级）
            string[] commonDevices = new string[]
            {
                "virtual-audio-capturer",  // Virtual Audio Cable
                "Stereo Mix (Realtek High Definition Audio)",  // Realtek 声卡
                "Stereo Mix",  // 通用立体声混音
                "What U Hear",  // Creative 声卡
                "Wave Out Mix",  // 某些声卡
                "Desktop Audio"  // 某些虚拟音频设备
            };
            
            // 尝试列出可用的音频设备
            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = _ffmpegPath,
                    Arguments = "-f dshow -list_devices true -i dummy",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                
                using (var process = Process.Start(processInfo))
                {
                    if (process != null)
                    {
                        process.WaitForExit(5000);
                        string output = process.StandardError.ReadToEnd();
                        
                        // 检查常见设备是否存在
                        foreach (var device in commonDevices)
                        {
                            if (output.Contains(device, StringComparison.OrdinalIgnoreCase))
                            {
                                WriteLine($"找到音频设备: {device}");
                                return device;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"检测音频设备失败", ex);
            }
            
            // 如果找不到，返回 null（将只录制视频）
            WriteWarning("未找到可用的音频设备");
            return string.Empty;
        }

        /// <summary>
        /// 构建无音频的 FFmpeg 命令
        /// </summary>
        private string BuildFfmpegCommandWithoutAudio(int videoBitrateKbps)
        {
            // 参考 Python 实现：使用 gdigrab 直接录制屏幕
            return $"-hide_banner -nostats -loglevel warning " +
                   $"-thread_queue_size 1024 " +
                   $"-f gdigrab " +
                   $"-framerate {_frameRate} " +
                   $"-offset_x {_offsetX} " +
                   $"-offset_y {_offsetY} " +
                   $"-video_size {_videoWidth}x{_videoHeight} " +
                   $"-use_wallclock_as_timestamps 1 " +
                   $"-i desktop " +
                   $"-c:v libx264 " +
                   $"-preset medium " +
                   $"-b:v {videoBitrateKbps}k " +
                   $"-pix_fmt yuv420p " +
                   $"-movflags +faststart " +
                   $"-r {_frameRate} " +
                   $"-g {_frameRate * 2} " +
                   $"-keyint_min {_frameRate} " +
                   $"-sc_threshold 0 " +
                   $"-threads 2 " +
                   $"-maxrate {videoBitrateKbps * 2}k " +
                   $"-bufsize {videoBitrateKbps * 4}k " +
                   $"-avoid_negative_ts make_zero " +
                   $"-fflags +genpts " +
                   $"-y \"{_outputPath}\"";
        }

        /// <summary>
        /// 设置音频文件路径（用于合并）
        /// </summary>
        public void SetAudioFile(string audioFilePath)
        {
            _tempAudioPath = audioFilePath;
            WriteLine($"设置音频文件路径: {audioFilePath}, 文件存在: {File.Exists(audioFilePath)}");
            if (File.Exists(audioFilePath))
            {
                var fileInfo = new FileInfo(audioFilePath);
                WriteLine($"音频文件大小: {fileInfo.Length} 字节");
            }
        }

        /// <summary>
        /// 请求停止录制（发送停止信号，但不等待完成）
        /// 用于与音频录制同时停止
        /// </summary>
        public void RequestStop()
        {
            if (!_isInitialized) return;

            try
            {
                Process? process = null;
                lock (_lockObject)
                {
                    process = _ffmpegProcess;
                }

                if (process != null && !process.HasExited && !_hasRequestedStop)
                {
                    // 发送 'q' 命令优雅退出（不等待）
                    try
                    {
                        process.StandardInput?.WriteLine("q");
                        process.StandardInput?.Flush();
                        _hasRequestedStop = true;
                        WriteLine($"已发送停止信号给 FFmpeg（等待完成中...）");
                    }
                    catch (Exception ex)
                    {
                        WriteWarning($"发送停止信号给 FFmpeg 失败: {ex.Message}");
                    }
                }
                else if (_hasRequestedStop)
                {
                    WriteLine($"停止信号已发送，等待 FFmpeg 完成...");
                }
            }
            catch (Exception ex)
            {
                WriteError($"请求停止录制失败", ex);
            }
        }

        /// <summary>
        /// 完成编码
        /// </summary>
        /// <param name="quickExit">快速退出模式，不等待 FFmpeg 完成（用于窗口关闭时）</param>
        public void Finish(bool quickExit = false)
        {
            if (!_isInitialized) return;

            try
            {
                Process? process = null;
                lock (_lockObject)
                {
                    process = _ffmpegProcess;
                }

                if (process != null)
                {
                    try
                    {
                        if (quickExit)
                        {
                            // 快速退出模式：立即终止 FFmpeg 进程
                            try
                            {
                                if (!process.HasExited)
                                {
                                    // 尝试优雅退出：发送 'q' 到标准输入
                                    try
                                    {
                                        process.StandardInput?.WriteLine("q");
                                        process.StandardInput?.Flush();
                                    }
                                    catch { }

                                    // 等待最多 2 秒
                                    if (!process.WaitForExit(2000))
                                    {
                                        process.Kill();
                                        process.WaitForExit(1000);
                                    }
                                }
                            }
                            catch { }
                            
                            try
                            {
                                process.Dispose();
                            }
                            catch { }
                        }
                        else
                        {
                            // 正常模式：优雅停止 FFmpeg
                            try
                            {
                                if (!process.HasExited)
                                {
                                    // 如果还没有发送停止信号，现在发送
                                    if (!_hasRequestedStop)
                                    {
                                        try
                                        {
                                            process.StandardInput?.WriteLine("q");
                                            process.StandardInput?.Flush();
                                            _hasRequestedStop = true;
                                            WriteLine($"已发送停止信号给 FFmpeg");
                                        }
                                        catch (Exception ex)
                                        {
                                            WriteWarning($"发送停止信号失败: {ex.Message}");
                                        }
                                    }
                                    else
                                    {
                                        WriteLine($"停止信号已发送，等待 FFmpeg 完成编码...");
                                    }

                                    // 等待 FFmpeg 完成（最多 30 秒）
                                    if (!process.WaitForExit(30000))
                                    {
                                        WriteWarning("FFmpeg 进程超时，强制终止");
                                        try
                                        {
                                            process.Kill();
                                            process.WaitForExit(5000);
                                        }
                                        catch { }
                                    }
                                    else
                                    {
                                        WriteLine($"FFmpeg 已正常退出");
                                    }
                                }
                                else
                                {
                                    WriteLine($"FFmpeg 进程已退出");
                                }
                            }
                            catch { }

                            // 检查输出文件
                            if (File.Exists(_outputPath))
                            {
                                var fileInfo = new FileInfo(_outputPath);
                                WriteLine($"视频文件已生成: {_outputPath}, 大小: {fileInfo.Length / 1024 / 1024} MB");
                                
                                if (fileInfo.Length == 0)
                                {
                                    WriteWarning($"视频文件大小为 0");
                                }
                                else
                                {
                                // 检查是否需要合并音频
                                if (_hasAudioInVideo)
                                {
                                    // 视频中已包含音频（FFmpeg 直接录制），无需合并
                                    WriteLine($"========== 录制完成 ==========");
                                    WriteLine($"✓ 最终文件: {_outputPath}");
                                    WriteLine($"✓ 生成方式: FFmpeg 直接生成 MP4（视频+音频同步录制）");
                                    WriteLine($"✓ 无需合并: 视频已包含音频流");
                                }
                                else
                                {
                                    // 检查音频文件路径
                                    WriteLine($"========== 准备合并音频 ==========");
                                    WriteLine($"检查音频文件路径: _tempAudioPath = {(string.IsNullOrEmpty(_tempAudioPath) ? "null" : _tempAudioPath)}");
                                    
                                    // 如果有音频文件，合并到 MP4
                                    if (!string.IsNullOrEmpty(_tempAudioPath))
                                    {
                                        if (File.Exists(_tempAudioPath))
                                        {
                                            WriteLine($"✓ 合并方式: FFmpeg 合并音频到 MP4");
                                            WriteLine($"  视频文件: {_outputPath}");
                                            WriteLine($"  音频文件: {_tempAudioPath}");
                                            WriteLine($"  说明: 使用 FFmpeg 将 NAudio 录制的 WAV 音频合并到视频 MP4");
                                            MergeAudioToMp4(_outputPath, _tempAudioPath);
                                        }
                                        else
                                        {
                                            WriteWarning($"音频文件不存在: {_tempAudioPath}");
                                        }
                                    }
                                    else
                                    {
                                        WriteWarning($"音频文件路径为空，跳过合并");
                                    }
                                }
                                }
                            }
                            else
                            {
                                WriteWarning($"视频文件未生成: {_outputPath}");
                            }
                        }
                    }
                    finally
                    {
                        try
                        {
                            process.Dispose();
                        }
                        catch { }
                    }
                }

                lock (_lockObject)
                {
                    _ffmpegProcess = null;
                }

                // 保留临时音频文件（用于调试）
                if (!string.IsNullOrEmpty(_tempAudioPath) && File.Exists(_tempAudioPath))
                {
                    WriteLine($"临时音频文件保留: {_tempAudioPath}");
                }
            }
            catch (Exception ex)
            {
                WriteError($"完成编码失败", ex);
            }
        }

        /// <summary>
        /// 合并音频到 MP4 文件
        /// </summary>
        private void MergeAudioToMp4(string videoPath, string audioPath)
        {
            try
            {
                if (!File.Exists(videoPath))
                {
                    WriteError($"视频文件不存在: {videoPath}");
                    return;
                }

                if (!File.Exists(audioPath))
                {
                    WriteError($"音频文件不存在: {audioPath}");
                    return;
                }

                var videoFileInfo = new FileInfo(videoPath);
                var audioFileInfo = new FileInfo(audioPath);

                WriteLine($"========== 开始合并音频到视频 ==========");
                WriteLine($"✓ 合并方式: FFmpeg 合并音频到 MP4");
                WriteLine($"  说明: 使用 FFmpeg 将 NAudio WasapiLoopbackCapture 录制的 WAV 音频合并到视频 MP4");
                WriteLine($"视频文件: {videoPath}");
                WriteLine($"  大小: {videoFileInfo.Length} 字节 ({videoFileInfo.Length / 1024 / 1024:F2} MB)");
                WriteLine($"  创建时间: {videoFileInfo.CreationTime:yyyy-MM-dd HH:mm:ss}");
                WriteLine($"  修改时间: {videoFileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");
                WriteLine($"音频文件: {audioPath}");
                WriteLine($"  大小: {audioFileInfo.Length} 字节 ({audioFileInfo.Length / 1024 / 1024:F2} MB)");
                WriteLine($"  创建时间: {audioFileInfo.CreationTime:yyyy-MM-dd HH:mm:ss}");
                WriteLine($"  修改时间: {audioFileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");
                WriteLine($"  说明: NAudio WasapiLoopbackCapture 录制（即使静音也会录制）");
                
                if (audioFileInfo.Length == 0)
                {
                    WriteWarning($"音频文件大小为 0，跳过合并");
                    return;
                }

                string tempOutput = Path.ChangeExtension(videoPath, ".temp.mp4");
                WriteLine($"临时输出文件: {tempOutput}");
                
                // 使用 FFmpeg 合并音频
                // 关键参数说明（确保音频正确合并，保留静音部分）：
                // -ss 0: 确保音频从开始读取（包括静音部分）
                // -itsoffset 0: 确保音频时间戳从 0 开始（必须在输入之前）
                // -c:v copy: 视频流直接复制，不重新编码
                // -c:a aac: 音频编码为 AAC
                // -b:a: 音频比特率
                // 注意：不指定 -ar，保持原始采样率（避免采样率转换导致音质下降和噪声）
                // -ac 2: 立体声
                // -async 1: 音频同步模式（与视频同步）
                // -avoid_negative_ts make_zero: 避免负时间戳，确保从 0 开始
                // -fflags +genpts: 生成时间戳，确保时间戳连续
                // -shortest: 以最短的流为准（视频或音频）
                string arguments = $"-i \"{videoPath}\" -ss 0 -itsoffset 0 -i \"{audioPath}\" -c:v copy -c:a aac -b:a {_config.AudioBitrate}k -ac 2 -map 0:v:0 -map 1:a:0 -shortest -async 1 -avoid_negative_ts make_zero -fflags +genpts -y \"{tempOutput}\"";
                
                WriteLine($"FFmpeg 路径: {_ffmpegPath}");
                WriteLine($"合并音频命令: {_ffmpegPath} {arguments}");
                WriteLine($"开始执行 FFmpeg 合并...");

                var processInfo = new ProcessStartInfo
                {
                    FileName = _ffmpegPath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = Process.Start(processInfo))
                {
                    if (process != null)
                    {
                        // 异步读取错误输出，避免缓冲区满导致进程阻塞
                        var errorOutputBuilder = new System.Text.StringBuilder();
                        var outputBuilder = new System.Text.StringBuilder();
                        
                        // 读取标准错误（FFmpeg 的主要输出）
                        Task errorReadTask = Task.Run(() =>
                        {
                            try
                            {
                                string? line;
                                while ((line = process.StandardError.ReadLine()) != null)
                                {
                                    errorOutputBuilder.AppendLine(line);
                                    if (line.Contains("Duration:") || line.Contains("Stream") || line.Contains("Audio") || line.Contains("Video") || line.Contains("error") || line.Contains("Error"))
                                    {
                                        WriteLine($"FFmpeg: {line}");
                                    }
                                }
                            }
                            catch (Exception ex) { WriteError($"读取 FFmpeg 错误输出时出错", ex); }
                        });
                        
                        Task outputReadTask = Task.Run(() =>
                        {
                            try
                            {
                                string? line;
                                while ((line = process.StandardOutput.ReadLine()) != null)
                                {
                                    outputBuilder.AppendLine(line);
                                }
                            }
                            catch (Exception ex) { WriteError($"读取 FFmpeg 标准输出时出错", ex); }
                        });

                        WriteLine($"等待 FFmpeg 进程完成（最多 120 秒）...");
                        bool exited = process.WaitForExit(120000); // 120秒超时
                        
                        Task.WaitAll(new[] { errorReadTask, outputReadTask }, TimeSpan.FromSeconds(5)); // 等待读取任务完成
                        
                        if (!exited)
                        {
                            WriteWarning("音频合并超时，强制终止");
                            string errorOutput = errorOutputBuilder.ToString();
                            string standardOutput = outputBuilder.ToString();
                            if (!string.IsNullOrEmpty(errorOutput)) { WriteLine($"FFmpeg 错误输出（超时前）: {errorOutput}"); }
                            if (!string.IsNullOrEmpty(standardOutput)) { WriteLine($"FFmpeg 标准输出（超时前）: {standardOutput}"); }
                            try { process.Kill(); process.WaitForExit(5000); } catch { }
                            try { if (File.Exists(tempOutput)) { File.Delete(tempOutput); } } catch { }
                        }
                        else if (process.ExitCode == 0)
                        {
                            WriteLine($"FFmpeg 进程成功退出（退出代码: 0）");
                            string errorOutput = errorOutputBuilder.ToString();
                            string standardOutput = outputBuilder.ToString();
                            
                            if (!string.IsNullOrEmpty(errorOutput))
                            {
                                WriteLine($"FFmpeg 完整错误输出: {errorOutput}");
                                if (errorOutput.Contains("Stream #1") && errorOutput.Contains("Audio")) { WriteLine($"✓ 检测到音频流信息"); }
                                else { WriteWarning($"⚠ 未在 FFmpeg 输出中检测到音频流信息"); }
                            }
                            if (!string.IsNullOrEmpty(standardOutput)) { WriteLine($"FFmpeg 完整标准输出: {standardOutput}"); }
                            
                            try
                            {
                                if (File.Exists(tempOutput))
                                {
                                    var tempFileInfo = new FileInfo(tempOutput);
                                    WriteLine($"合并后的文件大小: {tempFileInfo.Length} 字节 ({tempFileInfo.Length / 1024 / 1024:F2} MB)");
                                    WriteLine($"合并后的文件创建时间: {tempFileInfo.CreationTime:yyyy-MM-dd HH:mm:ss}");
                                    WriteLine($"合并后的文件修改时间: {tempFileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss}");
                                    
                                    // 使用 ffprobe 检查合并后的文件是否包含音频流
                                    try
                                    {
                                        string ffprobePath = _ffmpegPath.Replace("ffmpeg.exe", "ffprobe.exe");
                                        if (File.Exists(ffprobePath))
                                        {
                                            string probeArgs = $"-v error -select_streams a -show_entries stream=codec_name,codec_type,channels,sample_rate -of default=noprint_wrappers=1 \"{tempOutput}\"";
                                            var probeInfo = new ProcessStartInfo
                                            {
                                                FileName = ffprobePath,
                                                Arguments = probeArgs,
                                                UseShellExecute = false,
                                                CreateNoWindow = true,
                                                RedirectStandardOutput = true,
                                                RedirectStandardError = true
                                            };
                                            
                                            using (var probeProcess = Process.Start(probeInfo))
                                            {
                                                if (probeProcess != null)
                                                {
                                                    string probeOutput = probeProcess.StandardOutput.ReadToEnd();
                                                    string probeError = probeProcess.StandardError.ReadToEnd();
                                                    probeProcess.WaitForExit(5000);
                                                    
                                                    WriteLine($"ffprobe 检查合并后的文件:");
                                                    if (!string.IsNullOrEmpty(probeOutput))
                                                    {
                                                        WriteLine($"  音频流信息: {probeOutput.Trim()}");
                                                        if (probeOutput.Contains("codec_name=aac") || probeOutput.Contains("codec_type=audio")) { WriteLine($"✓ 确认合并后的文件包含音频流"); }
                                                        else { WriteWarning($"⚠ 合并后的文件可能不包含音频流"); }
                                                    }
                                                    if (!string.IsNullOrEmpty(probeError)) { WriteLine($"  ffprobe 错误: {probeError}"); }
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception probeEx) { WriteWarning($"无法使用 ffprobe 检查文件: {probeEx.Message}"); }
                                    
                                    // 备份原文件
                                    string backupPath = videoPath + ".backup";
                                    if (File.Exists(videoPath)) { File.Copy(videoPath, backupPath, true); WriteLine($"已备份原视频文件到: {backupPath}"); }
                                    
                                    File.Delete(videoPath);
                                    File.Move(tempOutput, videoPath);
                                    
                                    var finalFileInfo = new FileInfo(videoPath);
                                    WriteLine($"========== 合并完成 ==========");
                                    WriteLine($"✓ 最终文件: {videoPath}");
                                    WriteLine($"✓ 生成方式: FFmpeg 合并音频到 MP4");
                                    WriteLine($"  视频来源: FFmpeg gdigrab 录制");
                                    WriteLine($"  音频来源: NAudio WasapiLoopbackCapture 录制（即使静音也会录制）");
                                    WriteLine($"  合并工具: FFmpeg");
                                    WriteLine($"最终文件大小: {finalFileInfo.Length} 字节 ({finalFileInfo.Length / 1024 / 1024:F2} MB)");
                                }
                                else
                                {
                                    WriteWarning($"合并后的文件不存在: {tempOutput}");
                                    WriteLine($"FFmpeg 错误输出: {errorOutput}");
                                    WriteLine($"FFmpeg 标准输出: {standardOutput}");
                                }
                            }
                            catch (Exception ex) { WriteError($"替换文件失败", ex); }
                        }
                        else
                        {
                            try
                            {
                                string errorOutput = errorOutputBuilder.ToString();
                                string standardOutput = outputBuilder.ToString();
                                if (string.IsNullOrEmpty(errorOutput)) { try { errorOutput = process.StandardError.ReadToEnd(); } catch { } }
                                
                                WriteError($"音频合并失败，退出代码: {process.ExitCode}");
                                if (!string.IsNullOrEmpty(errorOutput)) { WriteLine($"FFmpeg 完整错误输出: {errorOutput}"); }
                                if (!string.IsNullOrEmpty(standardOutput)) { WriteLine($"FFmpeg 完整标准输出: {standardOutput}"); }
                            }
                            catch (Exception ex) { WriteError($"读取 FFmpeg 输出时出错", ex); }
                            try { if (File.Exists(tempOutput)) { File.Delete(tempOutput); } } catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"合并音频失败", ex);
            }
        }

        /// <summary>
        /// 查找 FFmpeg 可执行文件
        /// </summary>
        private string? FindFfmpeg()
        {
            // 首先检查程序目录
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string localPath = Path.Combine(exeDir, "ffmpeg.exe");
            if (File.Exists(localPath))
            {
                return localPath;
            }

            // 然后检查系统 PATH
            try
            {
                string? pathEnv = Environment.GetEnvironmentVariable("PATH");
                if (!string.IsNullOrEmpty(pathEnv))
                {
                    string[] paths = pathEnv.Split(Path.PathSeparator);
                    foreach (string path in paths)
                    {
                        string fullPath = Path.Combine(path, "ffmpeg.exe");
                        if (File.Exists(fullPath))
                        {
                            return fullPath;
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        public void Dispose()
        {
            try
            {
                Finish(quickExit: true);
            }
            catch { }
        }
    }
}
