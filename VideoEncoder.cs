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
        private int _videoWidth;      // 输出视频宽度（应用分辨率比例后）
        private int _videoHeight;     // 输出视频高度（应用分辨率比例后）
        private int _captureWidth;    // 捕获区域宽度（原始尺寸，不缩放）
        private int _captureHeight;   // 捕获区域高度（原始尺寸，不缩放）
        private int _frameRate;
        private int _audioSampleRate;
        private readonly object _lockObject = new object();
        private string? _ffmpegPath;
        private string? _tempAudioPath;
        private int _offsetX;
        private int _offsetY;
        private bool _hasAudioInVideo;     // 标记视频中是否已包含音频（FFmpeg 直接录制）
        private bool _hasRequestedStop;    // 标记是否已发送停止信号
        private Stream? _audioInputStream; // 音频管道流（用于实时合成）
        private int _audioBitsPerSample;   // 音频位深度（用于确定 FFmpeg 输入格式）
        private bool _audioIsFloat;        // 音频是否为浮点格式

        public VideoEncoder(string outputPath, RecordingConfig config)
        {
            _outputPath = outputPath;
            _config = config;
        }

        /// <summary>
        /// 初始化编码器
        /// </summary>
        /// <param name="videoWidth">输出视频宽度（应用分辨率比例后）</param>
        /// <param name="videoHeight">输出视频高度（应用分辨率比例后）</param>
        /// <param name="frameRate">帧率</param>
        /// <param name="audioSampleRate">音频采样率</param>
        /// <param name="audioChannels">音频通道数</param>
        /// <param name="offsetX">屏幕偏移 X（左上角 X 坐标）</param>
        /// <param name="offsetY">屏幕偏移 Y（左上角 Y 坐标）</param>
        /// <param name="captureWidth">捕获区域宽度（原始尺寸，不缩放，如果为0则使用videoWidth）</param>
        /// <param name="captureHeight">捕获区域高度（原始尺寸，不缩放，如果为0则使用videoHeight）</param>
        public void Initialize(
            int videoWidth,
            int videoHeight,
            int frameRate,
            int audioSampleRate,
            int audioChannels,
            int offsetX = 0,
            int offsetY = 0,
            int captureWidth = 0,
            int captureHeight = 0)
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

                // 如果未指定捕获尺寸，使用输出尺寸（100%分辨率时）
                _captureWidth = captureWidth > 0 ? captureWidth : videoWidth;
                _captureHeight = captureHeight > 0 ? captureHeight : videoHeight;

                // 确保宽度和高度是 2 的倍数（H.264 要求）
                _videoWidth = _videoWidth + (_videoWidth % 2);
                _videoHeight = _videoHeight + (_videoHeight % 2);
                _captureWidth = _captureWidth + (_captureWidth % 2);
                _captureHeight = _captureHeight + (_captureHeight % 2);

                // 查找 FFmpeg
                _ffmpegPath = FindFfmpeg();
                if (string.IsNullOrEmpty(_ffmpegPath))
                {
                    throw new FileNotFoundException("未找到 FFmpeg。请将 ffmpeg.exe 放在程序目录中，或添加到系统 PATH。");
                }

                // 不在这里设置 _tempAudioPath，等待 SetAudioFile 设置
                _tempAudioPath = null;
                _hasAudioInVideo = false;   // 初始化时假设没有音频
                _hasRequestedStop = false;  // 初始化时未发送停止信号
                _audioBitsPerSample = 16;   // 默认 16 位（NAudio WasapiLoopbackCapture 通常输出 16 位）
                _audioIsFloat = false;      // 默认整数格式（NAudio WasapiLoopbackCapture 通常输出整数 PCM）

                _isInitialized = true;

                if (_captureWidth != _videoWidth || _captureHeight != _videoHeight)
                {
                    WriteLine($"视频编码器初始化: 捕获尺寸 {_captureWidth}x{_captureHeight}, 输出尺寸 {_videoWidth}x{_videoHeight} @ {_frameRate}fps, 偏移: ({_offsetX}, {_offsetY})");
                }
                else
                {
                    WriteLine($"视频编码器初始化: {_videoWidth}x{_videoHeight} @ {_frameRate}fps, 偏移: ({_offsetX}, {_offsetY})");
                }
            }
            catch (Exception ex)
            {
                WriteError("初始化视频编码器失败", ex);
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
                string audioDevice = GetSystemAudioDevice();
                string arguments;

                if (!string.IsNullOrEmpty(audioDevice))
                {
                    // ✅ 同时录屏幕 + 系统音频
                    arguments = BuildFfmpegCommandWithAudio(videoBitrateKbps, audioDevice);
                    _hasAudioInVideo = true;
                    WriteLine("✓ 录制方式: FFmpeg 直接生成 MP4（视频+音频同步录制）");
                    WriteLine($"  音频设备: {audioDevice}");
                    WriteLine("  说明: 视频和音频在同一 FFmpeg 进程中录制，时间戳同步，无需后续合并");
                }
                else
                {
                    // 找不到音频设备，回退到只录制视频（后面用文件合并音频）
                    arguments = BuildFfmpegCommandWithoutAudio(videoBitrateKbps);
                    _hasAudioInVideo = false;
                    WriteLine("✓ 录制方式: 分步录制后合并（当前边录边合成暂时禁用）");
                    WriteLine("  视频: FFmpeg 录制（gdigrab）");
                    WriteLine("  音频: NAudio WasapiLoopbackCapture 录制（即使静音也会录制）");
                    WriteLine("  说明: 录制完成后使用 FFmpeg 合并音频到 MP4");
                }

                var processInfo = new ProcessStartInfo
                {
                    FileName = _ffmpegPath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardInput = true,  // 用于发送 'q' 停止
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                _ffmpegProcess = Process.Start(processInfo);
                if (_ffmpegProcess == null)
                {
                    throw new Exception("无法启动 FFmpeg 进程");
                }

                // 如果使用音频管道（当前逻辑 _hasAudioInVideo 表示“视频中有音频”，但我们这里用的是 dshow，不走 stdin 管道）
                // 保留逻辑以便以后启用 pipe:0 模式
                if (_hasAudioInVideo && _ffmpegProcess.StandardInput != null)
                {
                    _audioInputStream = _ffmpegProcess.StandardInput.BaseStream;
                    WriteLine("已准备音频管道（如启用 pipe:0 模式，可通过此流写入音频数据）");
                }

                // 异步读取 FFmpeg 错误输出，避免缓冲区堵塞
                Task.Run(() =>
                {
                    try
                    {
                        string? errorOutput = _ffmpegProcess.StandardError.ReadToEnd();
                        if (!string.IsNullOrEmpty(errorOutput))
                        {
                            if (errorOutput.Contains("error", StringComparison.OrdinalIgnoreCase) ||
                                errorOutput.Contains("failed", StringComparison.OrdinalIgnoreCase))
                            {
                                WriteLine($"FFmpeg 错误输出: {errorOutput}");
                            }
                        }
                    }
                    catch
                    {
                        // 忽略
                    }
                });

                WriteLine($"FFmpeg 录制已启动: {_outputPath}");
            }
            catch (Exception ex)
            {
                WriteError("启动录制失败", ex);
                throw;
            }
        }

        /// <summary>
        /// 构建带音频的 FFmpeg 命令（使用 dshow 直接录制系统音频）
        /// </summary>
        private string BuildFfmpegCommandWithAudio(int videoBitrateKbps, string audioDevice)
        {
            string scaleFilter = "";
            if (_captureWidth != _videoWidth || _captureHeight != _videoHeight)
            {
                scaleFilter = $"-vf scale={_videoWidth}:{_videoHeight} ";
            }

            return $"-hide_banner -nostats -loglevel warning " +
                   "-thread_queue_size 1024 " +
                   "-f gdigrab " +
                   $"-framerate {_frameRate} " +
                   $"-offset_x {_offsetX} " +
                   $"-offset_y {_offsetY} " +
                   $"-video_size {_captureWidth}x{_captureHeight} " +
                   "-use_wallclock_as_timestamps 1 " +
                   "-i desktop " +
                   "-thread_queue_size 1024 " +
                   "-f dshow " +
                   "-rtbufsize 256M " +
                   "-use_wallclock_as_timestamps 1 " +
                   $"-i \"audio={audioDevice}\" " +
                   scaleFilter +
                   "-c:v libx264 " +
                   "-preset medium " +
                   $"-b:v {videoBitrateKbps}k " +
                   "-pix_fmt yuv420p " +
                   "-c:a aac " +
                   $"-b:a {_config.AudioBitrate}k " +
                   "-movflags +faststart " +
                   $"-r {_frameRate} " +
                   $"-g {_frameRate * 2} " +
                   $"-keyint_min {_frameRate} " +
                   "-sc_threshold 0 " +
                   "-threads 2 " +
                   $"-maxrate {videoBitrateKbps * 2}k " +
                   $"-bufsize {videoBitrateKbps * 4}k " +
                   "-avoid_negative_ts make_zero " +
                   "-fflags +genpts " +
                   "-shortest " +
                   $"-y \"{_outputPath}\"";
        }

        /// <summary>
        /// 获取系统音频设备名称
        /// </summary>
        private string GetSystemAudioDevice()
        {
            string[] commonDevices = new string[]
            {
                "virtual-audio-capturer",
                "Stereo Mix (Realtek High Definition Audio)",
                "Stereo Mix",
                "What U Hear",
                "Wave Out Mix",
                "Desktop Audio"
            };

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
                WriteError("检测音频设备失败", ex);
            }

            WriteWarning("未找到可用的音频设备，将仅录制视频");
            return string.Empty;
        }

        /// <summary>
        /// 构建无音频的 FFmpeg 命令（只录制视频）
        /// </summary>
        private string BuildFfmpegCommandWithoutAudio(int videoBitrateKbps)
        {
            string scaleFilter = "";
            if (_captureWidth != _videoWidth || _captureHeight != _videoHeight)
            {
                scaleFilter = $"-vf scale={_videoWidth}:{_videoHeight} ";
            }

            return $"-hide_banner -nostats -loglevel warning " +
                   "-thread_queue_size 1024 " +
                   "-f gdigrab " +
                   $"-framerate {_frameRate} " +
                   $"-offset_x {_offsetX} " +
                   $"-offset_y {_offsetY} " +
                   $"-video_size {_captureWidth}x{_captureHeight} " +
                   "-use_wallclock_as_timestamps 1 " +
                   "-i desktop " +
                   scaleFilter +
                   "-c:v libx264 " +
                   "-preset medium " +
                   $"-b:v {videoBitrateKbps}k " +
                   "-pix_fmt yuv420p " +
                   "-movflags +faststart " +
                   $"-r {_frameRate} " +
                   $"-g {_frameRate * 2} " +
                   $"-keyint_min {_frameRate} " +
                   "-sc_threshold 0 " +
                   "-threads 2 " +
                   $"-maxrate {videoBitrateKbps * 2}k " +
                   $"-bufsize {videoBitrateKbps * 4}k " +
                   "-avoid_negative_ts make_zero " +
                   "-fflags +genpts " +
                   $"-y \"{_outputPath}\"";
        }

        /// <summary>
        /// （预留）构建带音频管道的 FFmpeg 命令（边录边合成）
        /// 当前未启用，仅保留以便后续升级
        /// </summary>
        private string BuildFfmpegCommandWithAudioPipe(int videoBitrateKbps)
        {
            string audioFormat;
            if (_audioIsFloat)
            {
                if (_audioBitsPerSample == 32)
                    audioFormat = "f32le";
                else if (_audioBitsPerSample == 64)
                    audioFormat = "f64le";
                else
                    audioFormat = "f32le";
            }
            else
            {
                if (_audioBitsPerSample == 16)
                    audioFormat = "s16le";
                else if (_audioBitsPerSample == 24)
                    audioFormat = "s24le";
                else if (_audioBitsPerSample == 32)
                    audioFormat = "s32le";
                else
                    audioFormat = "s16le";
            }

            WriteLine($"FFmpeg 音频管道格式: {audioFormat} ({_audioBitsPerSample}位, {(_audioIsFloat ? "浮点" : "整数")})");

            string scaleFilter = "";
            if (_captureWidth != _videoWidth || _captureHeight != _videoHeight)
            {
                scaleFilter = $"-vf scale={_videoWidth}:{_videoHeight} ";
            }

            return $"-hide_banner -nostats -loglevel warning " +
                   "-thread_queue_size 1024 " +
                   "-f gdigrab " +
                   $"-framerate {_frameRate} " +
                   $"-offset_x {_offsetX} " +
                   $"-offset_y {_offsetY} " +
                   $"-video_size {_captureWidth}x{_captureHeight} " +
                   "-use_wallclock_as_timestamps 1 " +
                   "-i desktop " +
                   $"-f {audioFormat} " +
                   $"-ar {_audioSampleRate} " +
                   "-ac 2 " +
                   "-use_wallclock_as_timestamps 1 " +
                   "-i pipe:0 " +
                   scaleFilter +
                   "-c:v libx264 " +
                   "-preset medium " +
                   $"-b:v {videoBitrateKbps}k " +
                   "-pix_fmt yuv420p " +
                   "-c:a aac " +
                   $"-b:a {_config.AudioBitrate}k " +
                   "-movflags +faststart " +
                   $"-r {_frameRate} " +
                   $"-g {_frameRate * 2} " +
                   $"-keyint_min {_frameRate} " +
                   "-sc_threshold 0 " +
                   "-threads 2 " +
                   $"-maxrate {videoBitrateKbps * 2}k " +
                   $"-bufsize {videoBitrateKbps * 4}k " +
                   "-avoid_negative_ts make_zero " +
                   "-fflags +genpts " +
                   "-shortest " +
                   $"-y \"{_outputPath}\"";
        }

        /// <summary>
        /// 设置音频文件路径（用于合并，仅在非实时合成模式下使用）
        /// </summary>
        public void SetAudioFile(string audioFilePath)
        {
            _tempAudioPath = audioFilePath;
            WriteLine($"设置音频文件路径（WAV文件）, 文件存在: {File.Exists(audioFilePath)}");
            if (File.Exists(audioFilePath))
            {
                var fileInfo = new FileInfo(audioFilePath);
                WriteLine($"音频文件大小: {fileInfo.Length} 字节");
            }
        }

        /// <summary>
        /// 写入音频数据到 FFmpeg 管道（用于实时合成，当前默认不启用）
        /// </summary>
        public void WriteAudioData(byte[] buffer, int bytesRecorded)
        {
            if (!_hasAudioInVideo || _audioInputStream == null || _ffmpegProcess == null || _ffmpegProcess.HasExited)
                return;

            try
            {
                lock (_lockObject)
                {
                    if (_audioInputStream != null && !_ffmpegProcess.HasExited)
                    {
                        _audioInputStream.Write(buffer, 0, bytesRecorded);
                        _audioInputStream.Flush();
                    }
                }
            }
            catch (Exception ex)
            {
                WriteWarning($"写入音频数据到 FFmpeg 管道失败: {ex.Message}");
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
                Process? process;
                lock (_lockObject)
                {
                    process = _ffmpegProcess;
                }

                if (process != null && !process.HasExited && !_hasRequestedStop)
                {
                    try
                    {
                        process.StandardInput?.WriteLine("q");
                        process.StandardInput?.Flush();
                        _hasRequestedStop = true;
                        WriteLine("已发送停止信号给 FFmpeg（等待完成中...）");
                    }
                    catch (Exception ex)
                    {
                        WriteWarning($"发送停止信号给 FFmpeg 失败: {ex.Message}");
                    }
                }
                else if (_hasRequestedStop)
                {
                    WriteLine("停止信号已发送，正在等待 FFmpeg 退出...");
                }
            }
            catch (Exception ex)
            {
                WriteError("请求停止录制失败", ex);
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
                Process? process;
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
                            // 快速退出模式：尽量优雅，必要时强杀
                            try
                            {
                                if (!process.HasExited)
                                {
                                    try
                                    {
                                        process.StandardInput?.WriteLine("q");
                                        process.StandardInput?.Flush();
                                    }
                                    catch { }

                                    if (!process.WaitForExit(2000))
                                    {
                                        process.Kill();
                                        process.WaitForExit(1000);
                                    }
                                }
                            }
                            catch { }
                        }
                        else
                        {
                            // 正常模式：给大文件足够长的时间收尾
                            try
                            {
                                if (!process.HasExited)
                                {
                                    if (!_hasRequestedStop)
                                    {
                                        try
                                        {
                                            process.StandardInput?.WriteLine("q");
                                            process.StandardInput?.Flush();
                                            _hasRequestedStop = true;
                                            WriteLine("已发送停止信号给 FFmpeg");
                                        }
                                        catch (Exception ex)
                                        {
                                            WriteWarning($"发送停止信号失败: {ex.Message}");
                                        }
                                    }
                                    else
                                    {
                                        WriteLine("停止信号已发送，等待 FFmpeg 完成编码...");
                                    }

                                    // ★★ 关键修改：录制结束后最多等 10 分钟，让 FFmpeg 有足够时间写完 2 小时以上的大文件
                                    const int MaxWaitForRecordExitMs = 10 * 60 * 1000; // 10 分钟
                                    if (!process.WaitForExit(MaxWaitForRecordExitMs))
                                    {
                                        WriteWarning($"FFmpeg 进程在 {MaxWaitForRecordExitMs / 1000} 秒内未退出，强制终止");
                                        try
                                        {
                                            process.Kill();
                                            process.WaitForExit(5000);
                                        }
                                        catch { }
                                    }
                                    else
                                    {
                                        WriteLine("FFmpeg 已正常退出");
                                    }
                                }
                                else
                                {
                                    WriteLine("FFmpeg 进程已提前退出");
                                }
                            }
                            catch
                            {
                                // 忽略
                            }

                            // 等待视频文件写入稳定
                            if (File.Exists(_outputPath))
                            {
                                WriteLine("等待视频文件完全写入...");
                                // 对大文件适当放宽等待时间
                                if (WaitForFileReady(_outputPath, maxWaitSeconds: 60))
                                {
                                    WriteLine("✓ 视频文件已完全写入");
                                }
                                else
                                {
                                    WriteWarning("等待视频文件写入超时，但继续处理");
                                }
                            }

                            // 检查输出文件
                            if (File.Exists(_outputPath))
                            {
                                var fileInfo = new FileInfo(_outputPath);
                                WriteLine($"MP4文件已生成, 大小: {fileInfo.Length / 1024 / 1024} MB");

                                if (fileInfo.Length == 0)
                                {
                                    WriteWarning("视频文件大小为 0");
                                }
                                else
                                {
                                    if (_hasAudioInVideo)
                                    {
                                        // 视频中已包含音频
                                        WriteLine("========== 录制完成 ==========");
                                        WriteLine($"✓ 最终文件: {_outputPath}");
                                        WriteLine("✓ 生成方式: FFmpeg 直接录制（视频 + 音频）");
                                        WriteLine("✓ 无需合并: 视频已包含音频流");
                                    }
                                    else
                                    {
                                        // 需要合并外部音频
                                        WriteLine("========== 准备合并音频 ==========");
                                        WriteLine($"检查音频文件路径: _tempAudioPath = {(string.IsNullOrEmpty(_tempAudioPath) ? "null" : _tempAudioPath)}");

                                        if (!string.IsNullOrEmpty(_tempAudioPath))
                                        {
                                            if (File.Exists(_tempAudioPath))
                                            {
                                                WriteLine("✓ 合并方式: FFmpeg 合并音频到 MP4");
                                                WriteLine($"  视频文件: {_outputPath}");
                                                WriteLine($"  音频文件: {_tempAudioPath}");
                                                WriteLine("  说明: 使用 FFmpeg 将 NAudio 录制的 WAV 音频合并到视频 MP4");
                                                MergeAudioToMp4(_outputPath, _tempAudioPath);
                                            }
                                            else
                                            {
                                                WriteWarning($"音频文件不存在: {_tempAudioPath}");
                                            }
                                        }
                                        else
                                        {
                                            WriteWarning("音频文件路径为空，跳过合并");
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
                        // 关闭音频管道流
                        try
                        {
                            if (_audioInputStream != null)
                            {
                                try
                                {
                                    _audioInputStream.Flush();
                                    _audioInputStream.Close();
                                }
                                catch { }

                                _audioInputStream = null;
                            }

                            process.Dispose();
                        }
                        catch { }
                    }
                }

                lock (_lockObject)
                {
                    _ffmpegProcess = null;
                }

                // 不删除临时音频文件，方便调试
            }
            catch (Exception ex)
            {
                WriteError("完成编码失败", ex);
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

                // 等待文件完全写入并验证文件完整性
                WriteLine("等待视频文件完全写入并验证完整性...");
                if (!WaitForFileReady(videoPath, maxWaitSeconds: 60))
                {
                    WriteWarning("等待视频文件写入超时，但继续尝试合并");
                }

                if (!ValidateMp4File(videoPath))
                {
                    WriteError("视频文件不完整或损坏（可能 moov 未写完），无法合并音频。");
                    return;
                }

                var videoFileInfo = new FileInfo(videoPath);
                var audioFileInfo = new FileInfo(audioPath);

                if (audioFileInfo.Length == 0)
                {
                    WriteWarning("音频文件大小为 0，跳过合并");
                    return;
                }

                WriteLine("正在合并音频到视频...");
                string tempOutput = Path.ChangeExtension(videoPath, ".temp.mp4");

                string arguments =
                    $"-i \"{videoPath}\" " +
                    "-ss 0 -itsoffset 0 " +
                    $"-i \"{audioPath}\" " +
                    "-c:v copy " +
                    "-c:a aac " +
                    $"-b:a {_config.AudioBitrate}k " +
                    "-ac 2 " +
                    "-map 0:v:0 -map 1:a:0 " +
                    "-shortest -async 1 " +
                    "-avoid_negative_ts make_zero " +
                    "-fflags +genpts " +
                    $"-y \"{tempOutput}\"";

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
                        var errorOutputBuilder = new System.Text.StringBuilder();
                        var outputBuilder = new System.Text.StringBuilder();

                        Task errorReadTask = Task.Run(() =>
                        {
                            try
                            {
                                string? line;
                                while ((line = process.StandardError.ReadLine()) != null)
                                {
                                    errorOutputBuilder.AppendLine(line);
                                    if (line.Contains("error", StringComparison.OrdinalIgnoreCase) ||
                                        line.Contains("failed", StringComparison.OrdinalIgnoreCase))
                                    {
                                        WriteLine($"FFmpeg: {line}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                WriteError("读取 FFmpeg 错误输出时出错", ex);
                            }
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
                            catch (Exception ex)
                            {
                                WriteError("读取 FFmpeg 标准输出时出错", ex);
                            }
                        });

                        // ★★ 关键修改：音频合并可能对 2 小时的视频耗时较长，把 120 秒改为 15 分钟
                        const int MaxWaitForMergeExitMs = 15 * 60 * 1000; // 15 分钟
                        WriteLine($"等待 FFmpeg 音频合并完成（最多 {MaxWaitForMergeExitMs / 1000} 秒）...");
                        bool exited = process.WaitForExit(MaxWaitForMergeExitMs);

                        Task.WaitAll(new[] { errorReadTask, outputReadTask }, TimeSpan.FromSeconds(5));

                        if (!exited)
                        {
                            WriteWarning("音频合并超时，强制终止");
                            string errorOutput = errorOutputBuilder.ToString();
                            string standardOutput = outputBuilder.ToString();
                            if (!string.IsNullOrEmpty(errorOutput))
                            {
                                WriteLine($"FFmpeg 错误输出（超时前）: {errorOutput}");
                            }
                            if (!string.IsNullOrEmpty(standardOutput))
                            {
                                WriteLine($"FFmpeg 标准输出（超时前）: {standardOutput}");
                            }
                            try
                            {
                                process.Kill();
                                process.WaitForExit(5000);
                            }
                            catch { }

                            try
                            {
                                if (File.Exists(tempOutput))
                                {
                                    File.Delete(tempOutput);
                                }
                            }
                            catch { }
                        }
                        else if (process.ExitCode == 0)
                        {
                            string errorOutput = errorOutputBuilder.ToString();
                            if (!string.IsNullOrEmpty(errorOutput) &&
                                (errorOutput.Contains("error", StringComparison.OrdinalIgnoreCase) ||
                                 errorOutput.Contains("failed", StringComparison.OrdinalIgnoreCase)))
                            {
                                WriteError($"FFmpeg 合并过程中出现错误: {errorOutput}");
                            }

                            try
                            {
                                if (File.Exists(tempOutput))
                                {
                                    if (File.Exists(videoPath))
                                    {
                                        File.Delete(videoPath);
                                    }
                                    File.Move(tempOutput, videoPath);

                                    var finalFileInfo = new FileInfo(videoPath);
                                    WriteLine($"✓ 音频合并完成，文件大小: {finalFileInfo.Length / 1024 / 1024:F2} MB");
                                }
                                else
                                {
                                    WriteWarning($"合并后的文件不存在: {tempOutput}");
                                    if (!string.IsNullOrEmpty(errorOutput))
                                    {
                                        WriteLine($"FFmpeg 错误输出: {errorOutput}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                WriteError("替换文件失败", ex);
                            }
                        }
                        else
                        {
                            try
                            {
                                string errorOutput = errorOutputBuilder.ToString();
                                string standardOutput = outputBuilder.ToString();
                                if (string.IsNullOrEmpty(errorOutput))
                                {
                                    try { errorOutput = process.StandardError.ReadToEnd(); } catch { }
                                }

                                WriteError($"音频合并失败，退出代码: {process.ExitCode}");
                                if (!string.IsNullOrEmpty(errorOutput))
                                {
                                    WriteLine($"FFmpeg 完整错误输出: {errorOutput}");
                                }
                                if (!string.IsNullOrEmpty(standardOutput))
                                {
                                    WriteLine($"FFmpeg 完整标准输出: {standardOutput}");
                                }
                            }
                            catch (Exception ex)
                            {
                                WriteError("读取 FFmpeg 输出时出错", ex);
                            }

                            try
                            {
                                if (File.Exists(tempOutput))
                                {
                                    File.Delete(tempOutput);
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError("合并音频失败", ex);
            }
        }

        /// <summary>
        /// 查找 FFmpeg 可执行文件
        /// </summary>
        private string? FindFfmpeg()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string localPath = Path.Combine(exeDir, "ffmpeg.exe");
            if (File.Exists(localPath))
            {
                return localPath;
            }

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

        /// <summary>
        /// 等待文件完全写入（通过检查文件大小是否稳定）
        /// </summary>
        private bool WaitForFileReady(string filePath, int maxWaitSeconds = 10)
        {
            if (!File.Exists(filePath))
                return false;

            int stableCount = 0;
            const int requiredStableChecks = 3;
            long lastSize = -1;
            int maxIterations = maxWaitSeconds * 10; // 每100ms检查一次
            int iteration = 0;

            while (iteration < maxIterations)
            {
                try
                {
                    var fileInfo = new FileInfo(filePath);
                    long currentSize = fileInfo.Length;

                    if (currentSize == lastSize)
                    {
                        stableCount++;
                        if (stableCount >= requiredStableChecks)
                        {
                            return true;
                        }
                    }
                    else
                    {
                        stableCount = 0;
                        lastSize = currentSize;
                    }
                }
                catch
                {
                    // 文件可能正在被写入，继续等待
                }

                Thread.Sleep(100);
                iteration++;
            }

            return false;
        }

        /// <summary>
        /// 验证 MP4 文件是否完整（检查 ftyp / moov 等基本结构）
        /// </summary>
        private bool ValidateMp4File(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            try
            {
                // 方法1：优先用 ffprobe 检查
                string? ffprobePath = FindFfprobe();
                if (!string.IsNullOrEmpty(ffprobePath))
                {
                    var processInfo = new ProcessStartInfo
                    {
                        FileName = ffprobePath,
                        Arguments = "-v error -select_streams v:0 -show_entries stream=codec_name -of default=noprint_wrappers=1 " +
                                   $"\"{filePath}\"",
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
                            if (process.ExitCode == 0)
                            {
                                return true;
                            }
                        }
                    }
                }

                // 方法2：简单检查 ftyp box
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    if (fs.Length < 32)
                        return false;

                    byte[] buffer = new byte[32];
                    int bytesRead = fs.Read(buffer, 0, 32);
                    if (bytesRead < 32)
                        return false;

                    string type = System.Text.Encoding.ASCII.GetString(buffer, 4, 4);
                    if (type == "ftyp")
                    {
                        return true;
                    }

                    return false;
                }
            }
            catch (Exception ex)
            {
                WriteWarning($"验证 MP4 文件时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 查找 FFprobe 可执行文件
        /// </summary>
        private string? FindFfprobe()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string localPath = Path.Combine(exeDir, "ffprobe.exe");
            if (File.Exists(localPath))
            {
                return localPath;
            }

            try
            {
                string? pathEnv = Environment.GetEnvironmentVariable("PATH");
                if (!string.IsNullOrEmpty(pathEnv))
                {
                    string[] paths = pathEnv.Split(Path.PathSeparator);
                    foreach (string path in paths)
                    {
                        string fullPath = Path.Combine(path, "ffprobe.exe");
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
            catch
            {
                // 忽略
            }
        }
    }
}
