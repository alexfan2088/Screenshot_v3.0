using System;
using System.Collections.Generic;
using System.IO;
using NAudio.Wave;
using NAudio.CoreAudioApi;
using NAudio.MediaFoundation;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 音频录制器（使用 NAudio 录制系统音频 - 声卡输出）
    /// - 统一输出为 16bit PCM，采样率和声道数由构造函数指定
    /// - 可选同时写入 WAV 文件（AudioOnly 模式），也可以只通过事件把音频送给 FFmpeg 管道
    /// - 即使过程静音，也会持续输出音频数据（静音部分为 0）
    /// </summary>
    public sealed class AudioRecorder : IDisposable
    {
        private readonly int _targetSampleRate;
        private readonly int _targetChannels;
        private readonly int _targetBitsPerSample;

        private WasapiLoopbackCapture? _loopback;
        private WaveFileWriter? _waveWriter;
        private string? _outputPath;

        private volatile bool _isRecording;
        private readonly object _lockObj = new();

        // 目标格式：统一使用 16bit PCM
        private readonly WaveFormat _targetFormat;
        // 捕获格式（可能是 32bit float 等）
        private WaveFormat? _captureFormat;
        
        // 用于重采样和声道转换的缓冲区
        private BufferedWaveProvider? _bufferedProvider;
        private MediaFoundationResampler? _resampler;

        /// <summary>
        /// 有新的 16bit PCM 音频数据时触发（buffer 长度 = bytesRecorded）
        /// </summary>
        public event Action<byte[], int>? AudioSampleAvailable;

        public AudioRecorder(int sampleRate, int channels, int bitsPerSample = 16)
        {
            if (sampleRate <= 0) throw new ArgumentOutOfRangeException(nameof(sampleRate));
            if (channels <= 0) throw new ArgumentOutOfRangeException(nameof(channels));
            if (bitsPerSample != 16)
                throw new ArgumentException("当前实现仅支持 16bit PCM 作为目标格式。", nameof(bitsPerSample));

            _targetSampleRate = sampleRate;
            _targetChannels = channels;
            _targetBitsPerSample = bitsPerSample;

            _targetFormat = new WaveFormat(_targetSampleRate, _targetBitsPerSample, _targetChannels);
        }

        /// <summary>
        /// 开始录制系统音频（管道模式，不写WAV文件，只通过事件传递）。
        /// </summary>
        public void StartPipeMode()
        {
            Start(null);
        }

        /// <summary>
        /// 开始录制系统音频。
        /// </summary>
        /// <param name="outputPath">
        /// 如果非空，则同时把音频写入该 WAV 文件（AudioOnly 模式）。
        /// 如果为 null，则只通过 AudioSampleAvailable 事件输出给视频编码器。
        /// </param>
        public void Start(string? outputPath = null)
        {
            lock (_lockObj)
            {
                if (_isRecording)
                {
                    WriteWarning("AudioRecorder.Start() 被重复调用，忽略。");
                    return;
                }

                try
                {
                    _outputPath = outputPath;

                    // 创建环回捕获（默认捕获系统播放的声音）
                    _loopback = new WasapiLoopbackCapture();
                    _captureFormat = _loopback.WaveFormat;

                    WriteLine($"[AudioRecorder] 捕获格式: {_captureFormat.SampleRate} Hz, {_captureFormat.BitsPerSample} bit, {_captureFormat.Channels} ch, {_captureFormat.Encoding}");
                    WriteLine($"[AudioRecorder] 目标格式: {_targetFormat.SampleRate} Hz, {_targetFormat.BitsPerSample} bit, {_targetFormat.Channels} ch");

                    // 如果需要重采样或声道转换，创建重采样器
                    bool needResample = _captureFormat.SampleRate != _targetFormat.SampleRate ||
                                       _captureFormat.Channels != _targetFormat.Channels ||
                                       _captureFormat.Encoding != WaveFormatEncoding.Pcm ||
                                       _captureFormat.BitsPerSample != 16;

                    if (needResample)
                    {
                        // 先转换为 16bit PCM（如果还不是）
                        WaveFormat intermediateFormat;
                        if (_captureFormat.Encoding == WaveFormatEncoding.IeeeFloat && _captureFormat.BitsPerSample == 32)
                        {
                            // 从 32bit float 转为 16bit PCM，保持原始采样率和声道数
                            intermediateFormat = new WaveFormat(_captureFormat.SampleRate, 16, _captureFormat.Channels);
                        }
                        else if (_captureFormat.Encoding == WaveFormatEncoding.Pcm && _captureFormat.BitsPerSample == 16)
                        {
                            // 已经是 16bit PCM
                            intermediateFormat = _captureFormat;
                        }
                        else
                        {
                            // 其他格式，尝试创建 16bit PCM 格式
                            intermediateFormat = new WaveFormat(_captureFormat.SampleRate, 16, _captureFormat.Channels);
                        }

                        // 创建缓冲提供者
                        _bufferedProvider = new BufferedWaveProvider(intermediateFormat)
                        {
                            BufferLength = 1024 * 1024, // 1MB 缓冲区
                            DiscardOnBufferOverflow = false
                        };

                        // 创建重采样器（从中间格式到目标格式）
                        _resampler = new MediaFoundationResampler(_bufferedProvider, _targetFormat);
                        WriteLine($"[AudioRecorder] 已创建重采样器: {intermediateFormat.SampleRate}Hz/{intermediateFormat.Channels}ch -> {_targetFormat.SampleRate}Hz/{_targetFormat.Channels}ch");
                    }

                    // 如果需要写入 WAV 文件，则创建 writer，目标格式为 16bit PCM
                    if (!string.IsNullOrEmpty(_outputPath))
                    {
                        _waveWriter = new WaveFileWriter(_outputPath, _targetFormat);
                        WriteLine($"[AudioRecorder] 将音频写入 WAV 文件: {_outputPath}");
                    }

                    _loopback.DataAvailable += OnDataAvailable;
                    _loopback.RecordingStopped += OnRecordingStopped;

                    _isRecording = true;
                    _loopback.StartRecording();
                    WriteLine("[AudioRecorder] 开始录制系统音频（WasapiLoopbackCapture）");
                }
                catch (Exception ex)
                {
                    WriteError("AudioRecorder.Start 失败", ex);
                    Cleanup();
                    throw;
                }
            }
        }

        /// <summary>
        /// 停止录制（同步返回，不会抛异常）。
        /// </summary>
        public void Stop()
        {
            lock (_lockObj)
            {
                if (!_isRecording)
                {
                    return;
                }

                try
                {
                    if (_loopback != null)
                    {
                        WriteLine("[AudioRecorder] 停止录制请求");
                        _loopback.StopRecording();
                    }
                }
                catch (Exception ex)
                {
                    // StopRecording 可能抛异常，但我们不让它影响主流程
                    WriteWarning($"AudioRecorder.Stop 调用 StopRecording 时异常: {ex.Message}");
                }
            }
        }

        private void OnDataAvailable(object? sender, WaveInEventArgs e)
        {
            try
            {
                if (!_isRecording || e.BytesRecorded <= 0)
                {
                    return;
                }

                if (_captureFormat == null)
                {
                    return;
                }

                // 把捕获到的数据统一转换为 16bit PCM（_targetFormat）
                byte[] pcm16 = ConvertToTargetFormat(e.Buffer, e.BytesRecorded);

                // 先写入 WAV 文件（如果有）
                if (_waveWriter != null && pcm16.Length > 0)
                {
                    _waveWriter.Write(pcm16, 0, pcm16.Length);
                    _waveWriter.Flush();
                }

                // 然后通知订阅者（VideoEncoder 通过管道写入 FFmpeg）
                if (pcm16.Length > 0)
                {
                    AudioSampleAvailable?.Invoke(pcm16, pcm16.Length);
                }
            }
            catch (Exception ex)
            {
                WriteError("AudioRecorder.OnDataAvailable 处理数据时异常", ex);
            }
        }

        private void OnRecordingStopped(object? sender, StoppedEventArgs e)
        {
            lock (_lockObj)
            {
                if (!_isRecording)
                {
                    return;
                }

                _isRecording = false;

                if (e.Exception != null)
                {
                    WriteWarning($"AudioRecorder.RecordingStopped 异常: {e.Exception.Message}");
                }

                Cleanup();
                WriteLine("[AudioRecorder] 录制已停止");
            }
        }

        /// <summary>
        /// 把捕获格式转换为目标格式（16bit PCM，目标采样率和声道数）。
        /// 支持格式转换、重采样和声道转换。
        /// </summary>
        private byte[] ConvertToTargetFormat(byte[] buffer, int bytesRecorded)
        {
            if (_captureFormat == null)
            {
                return Array.Empty<byte>();
            }

            // 如果不需要重采样，直接转换位宽
            if (_resampler == null)
            {
                return ConvertToPcm16Only(buffer, bytesRecorded, _captureFormat);
            }

            // 需要重采样：先转换为中间格式（16bit PCM，保持原始采样率和声道数）
            byte[] intermediatePcm = ConvertToPcm16Only(buffer, bytesRecorded, _captureFormat);
            
            if (intermediatePcm.Length == 0)
            {
                return Array.Empty<byte>();
            }

            // 将中间格式数据添加到缓冲区
            _bufferedProvider!.AddSamples(intermediatePcm, 0, intermediatePcm.Length);

            // 从重采样器读取转换后的数据
            // 计算期望的输出数据量（考虑采样率和声道数的变化）
            double sampleRateRatio = (double)_targetFormat.SampleRate / _bufferedProvider.WaveFormat.SampleRate;
            double channelRatio = (double)_targetFormat.Channels / _bufferedProvider.WaveFormat.Channels;
            int expectedOutputBytes = (int)(intermediatePcm.Length * sampleRateRatio * channelRatio);
            
            // 确保至少读取一些数据，但不超过缓冲区大小
            int readSize = Math.Min(expectedOutputBytes, _targetFormat.AverageBytesPerSecond / 4); // 250ms 缓冲区
            if (readSize < _targetFormat.BlockAlign)
            {
                readSize = _targetFormat.BlockAlign; // 至少一个音频块
            }
            
            byte[] readBuffer = new byte[readSize];
            int bytesRead = _resampler!.Read(readBuffer, 0, readBuffer.Length);
            
            if (bytesRead > 0)
            {
                byte[] result = new byte[bytesRead];
                Buffer.BlockCopy(readBuffer, 0, result, 0, bytesRead);
                return result;
            }

            return Array.Empty<byte>();
        }

        /// <summary>
        /// 只转换位宽和编码（32bit float -> 16bit PCM），不处理采样率和声道。
        /// </summary>
        private byte[] ConvertToPcm16Only(byte[] buffer, int bytesRecorded, WaveFormat captureFormat)
        {
            if (captureFormat.Encoding == WaveFormatEncoding.Pcm && captureFormat.BitsPerSample == 16)
            {
                // 已经是 16bit PCM，直接拷贝
                var result = new byte[bytesRecorded];
                Buffer.BlockCopy(buffer, 0, result, 0, bytesRecorded);
                return result;
            }

            if (captureFormat.Encoding == WaveFormatEncoding.IeeeFloat && captureFormat.BitsPerSample == 32)
            {
                int samples = bytesRecorded / 4; // 4 bytes per float
                var result = new byte[samples * 2]; // 2 bytes per short

                int outIndex = 0;
                for (int i = 0; i < samples; i++)
                {
                    float sample = BitConverter.ToSingle(buffer, i * 4);
                    // 限制在 [-1,1]
                    if (sample > 1.0f) sample = 1.0f;
                    if (sample < -1.0f) sample = -1.0f;
                    short int16 = (short)(sample * 32767f);
                    result[outIndex++] = (byte)(int16 & 0xFF);
                    result[outIndex++] = (byte)((int16 >> 8) & 0xFF);
                }

                return result;
            }

            // 其它格式简单降级处理：按 16bit PCM 直接截取（极少见）
            int sampleBytes = captureFormat.BitsPerSample / 8;
            if (sampleBytes <= 0)
            {
                // 不认识的格式，返回空
                return Array.Empty<byte>();
            }

            int totalSamples = bytesRecorded / sampleBytes;
            var fallback = new byte[totalSamples * 2];
            int dst = 0;
            for (int i = 0; i < totalSamples; i++)
            {
                int srcIndex = i * sampleBytes;
                short value = 0;
                if (sampleBytes >= 2)
                {
                    value = BitConverter.ToInt16(buffer, srcIndex);
                }
                fallback[dst++] = (byte)(value & 0xFF);
                fallback[dst++] = (byte)((value >> 8) & 0xFF);
            }

            return fallback;
        }

        private void Cleanup()
        {
            try
            {
                if (_resampler != null)
                {
                    _resampler.Dispose();
                    _resampler = null;
                }

                if (_bufferedProvider != null)
                {
                    _bufferedProvider.ClearBuffer();
                    _bufferedProvider = null;
                }

                if (_loopback != null)
                {
                    _loopback.DataAvailable -= OnDataAvailable;
                    _loopback.RecordingStopped -= OnRecordingStopped;
                    _loopback.Dispose();
                    _loopback = null;
                }

                if (_waveWriter != null)
                {
                    _waveWriter.Dispose();
                    _waveWriter = null;
                }
            }
            catch (Exception ex)
            {
                WriteWarning($"AudioRecorder.Cleanup 异常: {ex.Message}");
            }
        }

        public void Dispose()
        {
            lock (_lockObj)
            {
                if (_isRecording)
                {
                    try
                    {
                        _loopback?.StopRecording();
                    }
                    catch
                    {
                        // ignored
                    }
                    _isRecording = false;
                }

                Cleanup();
            }
        }
    }
}
