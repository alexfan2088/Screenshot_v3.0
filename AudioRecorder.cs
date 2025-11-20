using System;
using System.IO;
using NAudio.Wave;
using NAudio.Lame;
using NAudio.CoreAudioApi;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 音频录制器（使用 NAudio 录制系统音频 - 声卡输出）
    /// 保证：无论开始、过程还是结束是否有声音，都会生成与录制时长一致的音频，
    /// 静音部分用 0 填充。
    /// </summary>
    public sealed class AudioRecorder : IDisposable
    {
        private bool _isRecording;
        private string? _outputPath;
        private WasapiLoopbackCapture? _loopbackCapture;
        private WaveFileWriter? _waveFileWriter;
        private string? _tempWavPath;

        // === 时间轴 & 对齐参数 ===
        private int _bytesPerSecond;              // WaveFormat.AverageBytesPerSecond
        private int _blockAlign;                  // WaveFormat.BlockAlign（每帧字节数，必须对齐）
        private long _totalBytesWritten;          // 实际写入的总字节数
        private DateTime _recordingStartTime;     // Start 调用时间
        private DateTime _recordingStopTime;      // Stop 调用时间
        private DateTime _lastSampleTime;         // 当前已写样本对应的“时间轴”位置

        private DateTime _lastLogTime;
        private readonly object _dataLock = new object();

        public bool IsRecording => _isRecording;

        /// <summary>
        /// 音频样本回调（目前主要是给你预留扩展用）
        /// </summary>
        public event Action<byte[], int>? AudioSampleAvailable;

        /// <summary>
        /// 开始录制系统音频
        /// </summary>
        /// <param name="outAacPath">最终输出文件路径（这里只生成 wav，后续 FFmpeg 用 wav 合并）</param>
        public void Start(string outAacPath)
        {
            if (_isRecording) return;

            try
            {
                _outputPath = outAacPath;

                // 生成临时 WAV 路径
                var tempDir = Path.GetDirectoryName(outAacPath);
                var tempFileName = $"{Path.GetFileNameWithoutExtension(outAacPath)}.wav";
                _tempWavPath = Path.Combine(tempDir!, tempFileName);

                // 创建 WASAPI Loopback 捕获（系统播放输出）
                _loopbackCapture = new WasapiLoopbackCapture();

                // 创建 WAV 写入器
                _waveFileWriter = new WaveFileWriter(_tempWavPath, _loopbackCapture.WaveFormat);

                // 初始化时间轴 & 统计参数
                _bytesPerSecond = _loopbackCapture.WaveFormat.AverageBytesPerSecond;
                _blockAlign     = _loopbackCapture.WaveFormat.BlockAlign;   // ★ 关键：按这个对齐
                _totalBytesWritten   = 0;
                _recordingStartTime  = DateTime.Now;
                _lastSampleTime      = _recordingStartTime;
                _recordingStopTime   = DateTime.MinValue;
                _lastLogTime         = DateTime.Now;

                // 绑定事件
                _loopbackCapture.DataAvailable     += LoopbackCapture_DataAvailable;
                _loopbackCapture.RecordingStopped  += LoopbackCapture_RecordingStopped;

                // 开始录制
                _loopbackCapture.StartRecording();
                _isRecording = true;

                WriteLine($"开始录制系统音频到: {_tempWavPath}");
                WriteLine($"WaveFormat: {_loopbackCapture.WaveFormat}, BytesPerSecond = {_bytesPerSecond}, BlockAlign = {_blockAlign}");
                WriteLine("说明: 静音和空档部分会按 BlockAlign 对齐填充 0，保证后续解码正常。");
            }
            catch (Exception ex)
            {
                WriteError("开始录制系统音频失败", ex);
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

                // 记录停止时间（用于在 RecordingStopped 里补尾部静音）
                _recordingStopTime = DateTime.Now;

                _loopbackCapture?.StopRecording();
                WriteLine("停止录制系统音频，等待 RecordingStopped 回调补齐静音并关闭文件");
            }
            catch (Exception ex)
            {
                WriteError("停止录制系统音频失败", ex);
                throw;
            }
        }

        /// <summary>
        /// 数据到达事件：
        /// 1. 根据“当前时间 - _lastSampleTime”判断是否有空档期，有就补静音；
        /// 2. 写入本次数据；
        /// 3. 推进 _lastSampleTime 时间轴。
        /// </summary>
        private void LoopbackCapture_DataAvailable(object? sender, WaveInEventArgs e)
        {
            DateTime now = DateTime.Now;
            bool shouldLog = false;

            lock (_dataLock)
            {
                if (_waveFileWriter != null && _bytesPerSecond > 0 && _blockAlign > 0)
                {
                    // 1. 根据“真实时间轴”判断是否有空档
                    double gapSeconds = (now - _lastSampleTime).TotalSeconds;
                    if (gapSeconds < 0) gapSeconds = 0;

                    // 阈值可以适当放宽，避免因调度抖动引入过多补静音
                    const double GapThreshold = 0.10; // 100ms

                    if (gapSeconds > GapThreshold)
                    {
                        long missingBytes = (long)(gapSeconds * _bytesPerSecond);

                        // ★ 按 BlockAlign 对齐，避免“掰断采样帧”
                        missingBytes = AlignToBlock(missingBytes);

                        if (missingBytes > 0)
                        {
                            WriteSilenceBytes(missingBytes);
                            _totalBytesWritten += missingBytes;

                            double alignedSeconds = (double)missingBytes / _bytesPerSecond;
                            _lastSampleTime = _lastSampleTime.AddSeconds(alignedSeconds);

                            if (gapSeconds > 0.5)
                            {
                                WriteLine($"检测到空档 {gapSeconds:F3} 秒，对齐后补静音 {missingBytes} 字节（约 {alignedSeconds:F3} 秒），累计 {_totalBytesWritten} 字节");
                            }
                        }
                    }

                    // 2. 写入本次真实音频数据（NAudio 已保证 BytesRecorded 是 BlockAlign 的整数倍）
                    if (e.BytesRecorded > 0)
                    {
                        _waveFileWriter.Write(e.Buffer, 0, e.BytesRecorded);
                        _totalBytesWritten += e.BytesRecorded;

                        double dataSeconds = (double)e.BytesRecorded / _bytesPerSecond;
                        _lastSampleTime = _lastSampleTime.AddSeconds(dataSeconds);

                        AudioSampleAvailable?.Invoke(e.Buffer, e.BytesRecorded);
                    }

                    // 3. 每 5 秒打一条日志
                    if ((now - _lastLogTime).TotalSeconds >= 5)
                    {
                        shouldLog = true;
                        _lastLogTime = now;
                    }
                }
            }

            if (shouldLog)
            {
                WriteLine($"音频数据持续接收中... 当前包 {e.BytesRecorded} 字节，累计 {_totalBytesWritten} 字节");
            }
        }

        /// <summary>
        /// 把字节数向下对齐到 BlockAlign 的整数倍
        /// </summary>
        private long AlignToBlock(long bytes)
        {
            if (_blockAlign <= 0 || bytes <= 0) return 0;
            return (bytes / _blockAlign) * _blockAlign;
        }

        /// <summary>
        /// 写入指定长度的静音字节（全 0，且保证按 BlockAlign 对齐）
        /// </summary>
        private void WriteSilenceBytes(long bytesToWrite)
        {
            if (_waveFileWriter == null) return;

            bytesToWrite = AlignToBlock(bytesToWrite);
            if (bytesToWrite <= 0) return;

            const int BufferSize = 4096;
            byte[] buffer = new byte[BufferSize];

            while (bytesToWrite > 0)
            {
                int chunk = (int)Math.Min(bytesToWrite, BufferSize);
                // 这里 chunk 不一定是 BlockAlign 的整数倍，但前面 bytesToWrite 已经对齐，
                // 多次写入总和仍然是 BlockAlign 的倍数，WaveFileWriter 接受这种写法没问题。
                _waveFileWriter.Write(buffer, 0, chunk);
                bytesToWrite -= chunk;
            }
        }

        /// <summary>
        /// 录制停止事件：
        /// 保证最后“_lastSampleTime -> Stop 时间”这段也被静音填满。
        /// </summary>
        private void LoopbackCapture_RecordingStopped(object? sender, StoppedEventArgs e)
        {
            try
            {
                lock (_dataLock)
                {
                    if (_waveFileWriter != null && _bytesPerSecond > 0 && _blockAlign > 0)
                    {
                        if (_recordingStopTime == DateTime.MinValue)
                        {
                            _recordingStopTime = DateTime.Now;
                        }

                        double tailGapSeconds = (_recordingStopTime - _lastSampleTime).TotalSeconds;
                        if (tailGapSeconds < 0) tailGapSeconds = 0;

                        const double TailThreshold = 0.02; // 小于 20ms 忽略
                        if (tailGapSeconds > TailThreshold)
                        {
                            long missingBytes = (long)(tailGapSeconds * _bytesPerSecond);
                            missingBytes = AlignToBlock(missingBytes);

                            if (missingBytes > 0)
                            {
                                WriteSilenceBytes(missingBytes);
                                _totalBytesWritten += missingBytes;

                                double alignedSeconds = (double)missingBytes / _bytesPerSecond;
                                _lastSampleTime = _lastSampleTime.AddSeconds(alignedSeconds);

                                WriteLine($"录制结束补尾部静音 {tailGapSeconds:F3} 秒，对齐后 {alignedSeconds:F3} 秒，{missingBytes} 字节，最终总字节 {_totalBytesWritten}");
                            }
                        }
                    }
                }

                // 异常 / 正常日志
                if (e.Exception != null)
                {
                    WriteError("音频录制意外停止（异常）", e.Exception);
                }
                else
                {
                    WriteLine("音频录制正常停止");
                }

                // 关闭 WAV 文件
                _waveFileWriter?.Dispose();
                _waveFileWriter = null;

                if (!string.IsNullOrEmpty(_tempWavPath) && File.Exists(_tempWavPath))
                {
                    var fileInfo = new FileInfo(_tempWavPath);
                    WriteLine($"临时音频文件已保存: {_tempWavPath}, 大小: {fileInfo.Length} 字节");

                    if (_bytesPerSecond > 0)
                    {
                        double durationSeconds = (double)fileInfo.Length / _bytesPerSecond;
                        WriteLine($"音频文件时长(根据字节数计算): {durationSeconds:F2} 秒 ({durationSeconds / 60:F2} 分钟)");
                    }
                }

                if (_isRecording)
                {
                    WriteWarning("检测到录制意外停止，重置 _isRecording 标志");
                    _isRecording = false;
                }
            }
            catch (Exception ex)
            {
                WriteError("处理录制文件时出错", ex);
            }
            finally
            {
                _loopbackCapture?.Dispose();
                _loopbackCapture = null;
            }
        }

        /// <summary>
        /// 旧的 wav -> mp3 转换（目前流程用 FFmpeg，可以忽略）
        /// </summary>
        private void ConvertWavToM4A(string wavPath, string m4aPath)
        {
            try
            {
                using (var reader = new AudioFileReader(wavPath))
                using (var writer = new LameMP3FileWriter(m4aPath, reader.WaveFormat, 128))
                {
                    reader.CopyTo(writer);
                }
            }
            catch (Exception ex)
            {
                WriteError("转换音频格式失败", ex);
                // 如果转换失败，至少保留WAV文件
                File.Copy(wavPath, m4aPath, true);
            }
        }

        public void Dispose()
        {
            if (_isRecording) Stop();

            _loopbackCapture?.Dispose();
            _loopbackCapture = null;

            _waveFileWriter?.Dispose();
            _waveFileWriter = null;
        }
    }
}
