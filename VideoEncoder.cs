using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Text;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 负责调用 FFmpeg 进行屏幕录制和编码。
    /// 实现“视频 + 音频管道”一体录制：
    /// - 屏幕：使用 gdigrab 从桌面或指定区域采集
    /// - 音频：C# 侧使用 AudioRecorder 捕获系统声卡输出，经 NamedPipe 送入 FFmpeg
    /// - 不依赖 dshow / 虚拟声卡，不需要录制结束后的二次合成
    ///
    /// 使用方式（和你 MainWindow 里的逻辑匹配）：
    /// 1. var encoder = new VideoEncoder(outputPath, config);
    /// 2. encoder.Initialize(... 分辨率/帧率/区域/音频参数 ...);
    /// 3. encoder.SetAudioFormat(sampleRate, channels); // 比如 (48000, 2)
    /// 4. encoder.Start();
    /// 5. AudioRecorder.AudioSampleAvailable -> encoder.WriteAudioData(...);
    /// 6. 停止时：先停止 AudioRecorder，再调用 encoder.RequestStop()，最后 encoder.Finish()。
    /// </summary>
    public sealed class VideoEncoder : IDisposable
    {
        private readonly string _outputPath;
        private readonly RecordingConfig _config;

        private Process? _ffmpegProcess;
        private string? _ffmpegExePath;

        private int _outputWidth;
        private int _outputHeight;
        private int _frameRate;

        private int _captureWidth;
        private int _captureHeight;
        private int _offsetX;
        private int _offsetY;

        private int _audioSampleRate;
        private int _audioChannels;
        private int _audioBitsPerSample = 16; // 固定 16bit PCM

        private bool _useAudioPipe;

        // 音频 named pipe
        private NamedPipeServerStream? _audioPipeServer;
        private string? _audioPipeName;
        private readonly object _pipeLock = new();
        
        // 音频数据对齐缓冲区（16bit PCM 立体声需要4字节对齐）
        private byte[]? _audioAlignmentBuffer;
        private int _audioAlignmentBufferSize = 0;

        private volatile bool _hasRequestedStop;
        private volatile bool _hasStarted;
        private volatile bool _isDisposed;

        public VideoEncoder(string outputPath, RecordingConfig config)
        {
            _outputPath = outputPath ?? throw new ArgumentNullException(nameof(outputPath));
            _config = config ?? throw new ArgumentNullException(nameof(config));

            _ffmpegExePath = config.FfmpegPath;
            if (string.IsNullOrWhiteSpace(_ffmpegExePath))
            {
                throw new InvalidOperationException("RecordingConfig.FfmpegPath 未配置，无法启动 FFmpeg。");
            }
        }

        /// <summary>
        /// 初始化视频和音频参数（在 Start 之前调用）。
        /// </summary>
        public void Initialize(
            int outputWidth,
            int outputHeight,
            int frameRate,
            int audioSampleRate,
            int audioChannels,
            int offsetX,
            int offsetY,
            int captureWidth,
            int captureHeight)
        {
            _outputWidth = outputWidth;
            _outputHeight = outputHeight;
            _frameRate = frameRate;

            _audioSampleRate = audioSampleRate;
            _audioChannels = audioChannels;
            _audioBitsPerSample = 16; // 与 AudioRecorder 保持一致

            _offsetX = offsetX;
            _offsetY = offsetY;
            _captureWidth = captureWidth;
            _captureHeight = captureHeight;

            _useAudioPipe = _audioSampleRate > 0 && _audioChannels > 0;
        }

        /// <summary>
        /// 显式设置音频参数（兼容 MainWindow 中的 _videoEncoder.SetAudioFormat(_config.AudioSampleRate, 2) 调用）。
        /// </summary>
        public void SetAudioFormat(int sampleRate, int channels)
        {
            _audioSampleRate = sampleRate;
            _audioChannels = channels;
            _audioBitsPerSample = 16;
            _useAudioPipe = _audioSampleRate > 0 && _audioChannels > 0;

            WriteLine($"[VideoEncoder] SetAudioFormat: sampleRate={_audioSampleRate}, channels={_audioChannels}, bits={_audioBitsPerSample}");
        }

        /// <summary>
        /// 当前实现不再使用外部 WAV 文件合成，所以 SetAudioFile 保留空实现，仅记录日志以兼容旧调用。
        /// </summary>
        public void SetAudioFile(string? path)
        {
            if (!string.IsNullOrEmpty(path))
            {
                WriteLine($"[VideoEncoder] SetAudioFile 被调用，但当前版本使用音频管道实时合成，忽略外部音频文件: {path}");
            }
        }

        /// <summary>
        /// 启动 FFmpeg 进程，并在需要时创建 NamedPipe 等待音频连接。
        /// </summary>
        public void Start()
        {
            if (_hasStarted)
            {
                WriteWarning("[VideoEncoder] Start() 重复调用被忽略。");
                return;
            }

            if (string.IsNullOrEmpty(_ffmpegExePath) || !File.Exists(_ffmpegExePath))
            {
                throw new FileNotFoundException("找不到 FFmpeg 可执行文件。", _ffmpegExePath);
            }

            try
            {
                _hasStarted = true;
                _hasRequestedStop = false;

                // 如果需要音频，则先准备 NamedPipe
                if (_useAudioPipe)
                {
                    _audioPipeName = $"screenshot_audio_{Guid.NewGuid():N}";
                    // FFmpeg 访问路径形式：\\.\pipe\name
                    string fullPipePath = $@"\\.\pipe\{_audioPipeName}";

                    _audioPipeServer = new NamedPipeServerStream(
                        _audioPipeName,
                        PipeDirection.Out,
                        1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);

                    WriteLine($"[VideoEncoder] 已创建音频 NamedPipe: {fullPipePath}");

                    // 启动 FFmpeg
                    string arguments = BuildFfmpegCommandWithAudioPipe(fullPipePath);
                    StartFfmpegProcess(arguments);

                    // 等待 FFmpeg 连接到 NamedPipe（异步等待）
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            WriteLine("[VideoEncoder] 等待 FFmpeg 连接音频管道...");
                            _audioPipeServer!.WaitForConnection();
                            WriteLine("[VideoEncoder] 音频管道连接成功（FFmpeg 已连接）。");
                        }
                        catch (Exception ex)
                        {
                            WriteError("VideoEncoder 等待音频管道连接时异常", ex);
                        }
                    });
                }
                else
                {
                    // 纯视频模式
                    string arguments = BuildFfmpegCommandWithoutAudio();
                    StartFfmpegProcess(arguments);
                }
            }
            catch (Exception ex)
            {
                WriteError("VideoEncoder.Start 启动 FFmpeg 失败", ex);
                throw;
            }
        }

        private void StartFfmpegProcess(string arguments)
        {
            var psi = new ProcessStartInfo
            {
                FileName = _ffmpegExePath!,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardInput = true, // 允许通过标准输入发送 'q' 来停止
                RedirectStandardError = true,
                RedirectStandardOutput = false,
                StandardErrorEncoding = Encoding.UTF8
            };

            _ffmpegProcess = new Process { StartInfo = psi, EnableRaisingEvents = true };
            _ffmpegProcess.ErrorDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    WriteFfmpegLog(e.Data);
                }
            };

            WriteLine($"[VideoEncoder] 启动 FFmpeg: {_ffmpegExePath} {arguments}");
            bool started = _ffmpegProcess.Start();
            if (!started)
            {
                throw new InvalidOperationException("FFmpeg 进程启动失败。");
            }

            _ffmpegProcess.BeginErrorReadLine();
        }

        private string BuildFfmpegCommandWithoutAudio()
        {
            // 只录制视频（没有音频）
            // 使用 gdigrab 采集屏幕指定区域
            var sb = new StringBuilder();

            sb.Append($"-f gdigrab -framerate {_frameRate} ");

            if (_captureWidth > 0 && _captureHeight > 0)
            {
                sb.Append($"-offset_x {_offsetX} -offset_y {_offsetY} ");
                sb.Append($"-video_size {_captureWidth}x{_captureHeight} ");
            }

            sb.Append("-draw_mouse 1 -i desktop ");

            // 视频编码参数
            sb.Append("-c:v libx264 ");
            sb.Append($"-preset {(_config.FfmpegPreset ?? "veryfast")} ");
            sb.Append($"-crf {_config.Crf} ");
            sb.Append("-pix_fmt yuv420p ");

            // 如果需要缩放到输出分辨率，可以加 scale 滤镜
            if (_outputWidth > 0 && _outputHeight > 0 &&
                (_outputWidth != _captureWidth || _outputHeight != _captureHeight))
            {
                sb.Append($"-vf scale={_outputWidth}:{_outputHeight} ");
            }

            // 直接写出 MP4 文件
            sb.Append("-movflags +faststart ");
            sb.Append($"\"{_outputPath}\"");

            return sb.ToString();
        }

        private string BuildFfmpegCommandWithAudioPipe(string pipePath)
        {
            var sb = new StringBuilder();

            // 屏幕输入（gdigrab）
            // 增加输入缓冲区大小，避免 gdigrab 初始化时丢帧
            sb.Append($"-thread_queue_size 512 ");
            sb.Append($"-f gdigrab -framerate {_frameRate} ");
            if (_captureWidth > 0 && _captureHeight > 0)
            {
                sb.Append($"-offset_x {_offsetX} -offset_y {_offsetY} ");
                sb.Append($"-video_size {_captureWidth}x{_captureHeight} ");
            }
            sb.Append("-draw_mouse 1 -i desktop ");

            // 音频输入（NamedPipe，16bit PCM）
            // 增加音频输入缓冲区大小
            sb.Append($"-thread_queue_size 512 ");
            sb.Append($"-f s16le -ar {_audioSampleRate} -ac {_audioChannels} -i \"{pipePath}\" ");

            // 编码设置
            sb.Append("-c:v libx264 ");
            sb.Append($"-preset {(_config.FfmpegPreset ?? "veryfast")} ");
            sb.Append($"-crf {_config.Crf} ");
            sb.Append("-pix_fmt yuv420p ");

            sb.Append("-c:a aac -b:a 192k ");

            // 不使用 -shortest，因为如果音频管道关闭时视频流还没开始，会导致没有视频帧
            // 改为通过 RequestStop() 发送 'q' 来优雅停止 FFmpeg
            // 这样可以确保视频流有足够时间开始和编码

            // 为了让停止后几乎立即可播放，使用 faststart
            sb.Append("-movflags +faststart ");

            sb.Append($"\"{_outputPath}\"");

            return sb.ToString();
        }

        /// <summary>
        /// 把 16bit PCM 音频数据写入 FFmpeg 的 NamedPipe。
        /// 确保数据按 BlockAlign（4字节，16bit立体声）对齐。
        /// </summary>
        public void WriteAudioData(byte[] buffer, int count)
        {
            if (!_useAudioPipe || buffer == null || count <= 0)
            {
                return;
            }

            try
            {
                lock (_pipeLock)
                {
                    if (_audioPipeServer == null || !_audioPipeServer.IsConnected)
                    {
                        // 连接尚未建立，直接丢弃这部分音频，避免阻塞
                        return;
                    }

                    // 16bit PCM 立体声需要4字节对齐（2字节/样本 × 2声道）
                    const int blockAlign = 4;
                    
                    // 初始化对齐缓冲区
                    if (_audioAlignmentBuffer == null)
                    {
                        _audioAlignmentBuffer = new byte[blockAlign];
                        _audioAlignmentBufferSize = 0;
                    }

                    // 将新数据与缓冲区中的数据合并
                    int totalBytes = _audioAlignmentBufferSize + count;
                    byte[] combinedBuffer = new byte[totalBytes];
                    
                    if (_audioAlignmentBufferSize > 0)
                    {
                        Buffer.BlockCopy(_audioAlignmentBuffer, 0, combinedBuffer, 0, _audioAlignmentBufferSize);
                    }
                    Buffer.BlockCopy(buffer, 0, combinedBuffer, _audioAlignmentBufferSize, count);

                    // 计算对齐后的数据量
                    int alignedBytes = (totalBytes / blockAlign) * blockAlign;
                    int remainder = totalBytes % blockAlign;

                    // 写入对齐后的数据
                    if (alignedBytes > 0)
                    {
                        _audioPipeServer.Write(combinedBuffer, 0, alignedBytes);
                        _audioPipeServer.Flush();
                    }

                    // 保存未对齐的剩余数据到缓冲区
                    if (remainder > 0)
                    {
                        Buffer.BlockCopy(combinedBuffer, alignedBytes, _audioAlignmentBuffer, 0, remainder);
                        _audioAlignmentBufferSize = remainder;
                    }
                    else
                    {
                        _audioAlignmentBufferSize = 0;
                    }
                }
            }
            catch (ObjectDisposedException)
            {
                // 停止时可能已经 Dispose，忽略
            }
            catch (IOException ioEx)
            {
                // FFmpeg 退出后写入会抛 IO 异常，忽略
                WriteWarning($"WriteAudioData 写入管道时 IO 异常: {ioEx.Message}");
            }
            catch (Exception ex)
            {
                WriteWarning($"WriteAudioData 写入管道时异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 请求停止录制。
        /// 对于管道模式：先关闭音频管道，然后发送 'q' 给 FFmpeg 标准输入来优雅停止。
        /// 对于纯视频模式：发送 'q' 来优雅停止，最后必要时 Kill。
        /// </summary>
        public void RequestStop()
        {
            if (_hasRequestedStop)
            {
                return;
            }

            _hasRequestedStop = true;

            try
            {
                if (_useAudioPipe)
                {
                    // 先处理对齐缓冲区中的剩余数据（如果有）
                    lock (_pipeLock)
                    {
                        if (_audioPipeServer != null && _audioPipeServer.IsConnected)
                        {
                            try
                            {
                                // 如果有未对齐的剩余数据，补齐到4字节（用0填充）
                                if (_audioAlignmentBufferSize > 0)
                                {
                                    const int blockAlign = 4;
                                    int padding = blockAlign - _audioAlignmentBufferSize;
                                    if (padding > 0 && padding < blockAlign)
                                    {
                                        // 用0填充到对齐边界
                                        byte[] padded = new byte[blockAlign];
                                        Buffer.BlockCopy(_audioAlignmentBuffer!, 0, padded, 0, _audioAlignmentBufferSize);
                                        // 剩余字节已经是0（新数组默认值）
                                        _audioPipeServer.Write(padded, 0, blockAlign);
                                    }
                                    else if (_audioAlignmentBufferSize == blockAlign)
                                    {
                                        // 正好对齐，直接写入
                                        _audioPipeServer.Write(_audioAlignmentBuffer!, 0, blockAlign);
                                    }
                                    _audioAlignmentBufferSize = 0;
                                    _audioPipeServer.Flush();
                                }
                            }
                            catch
                            {
                                // ignore
                            }
                        }

                        // 关闭音频管道（发送 EOF）
                        if (_audioPipeServer != null)
                        {
                            try
                            {
                                if (_audioPipeServer.IsConnected)
                                {
                                    _audioPipeServer.Flush();
                                }
                            }
                            catch
                            {
                                // ignore
                            }

                            _audioPipeServer.Dispose();
                            _audioPipeServer = null;
                            WriteLine("[VideoEncoder] 已关闭音频 NamedPipe");
                        }
                    }

                    // 等待一小段时间，确保 FFmpeg 处理完剩余的音频数据
                    System.Threading.Thread.Sleep(100);
                }

                // 发送 'q' 给 FFmpeg 标准输入来优雅停止
                if (_ffmpegProcess != null && !_ffmpegProcess.HasExited)
                {
                    try
                    {
                        if (_ffmpegProcess.StartInfo.RedirectStandardInput && _ffmpegProcess.StandardInput != null)
                        {
                            _ffmpegProcess.StandardInput.Write('q');
                            _ffmpegProcess.StandardInput.Flush();
                            WriteLine("[VideoEncoder] 已发送停止信号 'q' 给 FFmpeg");
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteWarning($"VideoEncoder.RequestStop 发送 'q' 时异常: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteWarning($"VideoEncoder.RequestStop 异常: {ex.Message}");
            }
        }

        private void TryKillFfmpegGracefully()
        {
            try
            {
                if (_ffmpegProcess == null || _ffmpegProcess.HasExited)
                {
                    return;
                }

                // 先给一点时间让它自己退出
                if (!_ffmpegProcess.WaitForExit(1000))
                {
                    WriteWarning("[VideoEncoder] FFmpeg 在 1 秒内未退出，调用 Kill()");
                    _ffmpegProcess.Kill(true);
                }
            }
            catch (Exception ex)
            {
                WriteWarning($"TryKillFfmpegGracefully 异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 等待 FFmpeg 完成并清理资源。
        /// </summary>
        public void Finish()
        {
            try
            {
                if (_ffmpegProcess != null)
                {
                    try
                    {
                        if (!_ffmpegProcess.HasExited)
                        {
                            // 最长等待 30 秒
                            if (!_ffmpegProcess.WaitForExit(30000))
                            {
                                WriteWarning("[VideoEncoder] FFmpeg 在 30 秒内仍未退出，Kill()");
                                _ffmpegProcess.Kill(true);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteWarning($"VideoEncoder.Finish 等待 FFmpeg 退出时异常: {ex.Message}");
                    }
                    finally
                    {
                        // 关闭标准输入流
                        try
                        {
                            if (_ffmpegProcess.StartInfo.RedirectStandardInput && _ffmpegProcess.StandardInput != null)
                            {
                                _ffmpegProcess.StandardInput.Close();
                            }
                        }
                        catch
                        {
                            // ignore
                        }

                        _ffmpegProcess.Dispose();
                        _ffmpegProcess = null;
                    }
                }

                lock (_pipeLock)
                {
                    if (_audioPipeServer != null)
                    {
                        _audioPipeServer.Dispose();
                        _audioPipeServer = null;
                    }
                }
            }
            catch (Exception ex)
            {
                WriteWarning($"VideoEncoder.Finish 清理资源时异常: {ex.Message}");
            }
        }

        private static void WriteFfmpegLog(string line)
        {
            WriteLine("[FFmpeg] " + line);
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;

            try
            {
                RequestStop();
                Finish();
            }
            catch
            {
                // ignored
            }
        }
    }
}
