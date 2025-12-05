using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    /// <summary>
    /// 录制配置数据模型
    /// </summary>
    public class RecordingConfig
    {
        // 视频参数
        public int VideoResolutionScale { get; set; } = 100; // 分辨率比例 10-100 (100 = 整个屏幕)
        public int VideoFrameRate { get; set; } = 60; // 帧率 15/24/30/60 (60 = 最优画质)
        public string VideoBitrate { get; set; } = "High"; // 码率: Low/Medium/High/Auto (High = 最优画质)

        // 音频参数
        public int AudioSampleRate { get; set; } = 44100; // 采样率: 22050/44100
        public int AudioBitrate { get; set; } = 192; // 音频比特率: 64/128/192 kbps (192 = 最优画质)
        
        // 日志参数
        public int LogEnabled { get; set; } = 1; // 日志开关 1=启用，0=禁用
        public int LogFileMode { get; set; } = 0; // 日志文件模式 0=覆盖，1=叠加

        // 自定义录制区域
        public bool UseCustomRegion { get; set; } = false;
        public int RegionLeft { get; set; }
        public int RegionTop { get; set; }
        public int RegionWidth { get; set; }
        public int RegionHeight { get; set; }

        // 截图参数
        public double ScreenChangeRate { get; set; } = 11.12; // 屏幕变化率 1-1000%，默认11.12%
        public int ScreenshotInterval { get; set; } = 10; // 截图间隔 1-65535秒，默认10秒
        
        // PPT和PDF生成参数
        public bool GeneratePPT { get; set; } = true; // 生成PPT，默认启用
        public bool GeneratePDF { get; set; } = false; // 生成PDF，默认禁用
        
        // 截图参数
        public bool KeepJpgFiles { get; set; } = true; // 保留JPG文件，默认启用
        
        // 区域显示参数
        public bool ShowRegionOverlay { get; set; } = false; // 显示上次记录的矩形框，默认不显示

        // FFmpeg参数
        public string? FfmpegPath { get; set; } // FFmpeg可执行文件路径，如果为空则自动查找
        public string? FfmpegPreset { get; set; } = "veryfast"; // FFmpeg编码预设: ultrafast/superfast/veryfast/faster/fast/medium/slow/slower/veryslow
        public int Crf { get; set; } = 23; // CRF值: 0-51，默认23（18-28为常用范围，值越小质量越高）

        /// <summary>
        /// 获取视频码率（Mbps）
        /// </summary>
        /// <param name="outputWidth">输出视频宽度（应用分辨率比例后），如果为0则使用默认值</param>
        /// <param name="outputHeight">输出视频高度（应用分辨率比例后），如果为0则使用默认值</param>
        public int GetVideoBitrateMbps(int outputWidth = 0, int outputHeight = 0)
        {
            return VideoBitrate switch
            {
                "Low" => 1,
                "Medium" => 3,
                "High" => 5,
                "Auto" => GetAutoVideoBitrate(outputWidth, outputHeight),
                _ => 3
            };
        }

        /// <summary>
        /// 根据分辨率自动计算码率
        /// </summary>
        /// <param name="outputWidth">输出视频宽度（应用分辨率比例后）</param>
        /// <param name="outputHeight">输出视频高度（应用分辨率比例后）</param>
        private int GetAutoVideoBitrate(int outputWidth = 0, int outputHeight = 0)
        {
            // 如果提供了实际分辨率，使用实际分辨率；否则使用默认值
            int scaledWidth, scaledHeight;
            if (outputWidth > 0 && outputHeight > 0)
            {
                scaledWidth = outputWidth;
                scaledHeight = outputHeight;
            }
            else
            {
                // 假设原始分辨率为 1920x1080
                scaledWidth = (int)(1920 * VideoResolutionScale / 100.0);
                scaledHeight = (int)(1080 * VideoResolutionScale / 100.0);
            }
            
            int pixels = scaledWidth * scaledHeight;

            // 根据像素数估算码率
            if (pixels < 500000) return 1;      // < 720p
            if (pixels < 1000000) return 2;     // 720p
            return 3;                            // >= 1080p
        }

        /// <summary>
        /// 计算预计文件大小（MB/分钟）
        /// </summary>
        /// <param name="outputWidth">输出视频宽度（应用分辨率比例后），如果为0则使用默认值</param>
        /// <param name="outputHeight">输出视频高度（应用分辨率比例后），如果为0则使用默认值</param>
        public double EstimateFileSizePerMinute(int outputWidth = 0, int outputHeight = 0)
        {
            // 使用实际分辨率计算码率（如果码率是Auto，会根据分辨率自动调整）
            int videoBitrateMbps = GetVideoBitrateMbps(outputWidth, outputHeight);
            double audioBitrateMbps = AudioBitrate / 1000.0; // kbps to Mbps
            return (videoBitrateMbps + audioBitrateMbps) * 60.0 / 8.0; // MB/分钟
        }

        /// <summary>
        /// 保存配置到文件
        /// </summary>
        public void Save(string configPath)
        {
            try
            {
                string json = JsonConvert.SerializeObject(this, Formatting.Indented);
                File.WriteAllText(configPath, json);
            }
            catch (Exception ex)
            {
                WriteError($"保存配置失败", ex);
            }
        }

        /// <summary>
        /// 从文件加载配置
        /// </summary>
        public static RecordingConfig Load(string configPath)
        {
            try
            {
                RecordingConfig config;
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var deserialized = JsonConvert.DeserializeObject<RecordingConfig>(json);
                    config = deserialized ?? new RecordingConfig();
                }
                else
                {
                    config = new RecordingConfig();
                }

                // 如果FFmpeg路径未配置，尝试自动查找
                if (string.IsNullOrWhiteSpace(config.FfmpegPath))
                {
                    config.FfmpegPath = FindFfmpegPath();
                }

                return config;
            }
            catch (Exception ex)
            {
                WriteError($"加载配置失败", ex);
            }
            var defaultConfig = new RecordingConfig();
            defaultConfig.FfmpegPath = FindFfmpegPath();
            return defaultConfig;
        }

        /// <summary>
        /// 自动查找FFmpeg可执行文件路径
        /// </summary>
        private static string? FindFfmpegPath()
        {
            // 常见路径列表
            var possiblePaths = new[]
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ffmpeg.exe"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ffmpeg", "ffmpeg.exe"),
                @"C:\ffmpeg\bin\ffmpeg.exe",
                @"C:\Program Files\ffmpeg\bin\ffmpeg.exe",
                @"C:\Program Files (x86)\ffmpeg\bin\ffmpeg.exe",
            };

            // 检查PATH环境变量
            var pathEnv = Environment.GetEnvironmentVariable("PATH");
            if (!string.IsNullOrEmpty(pathEnv))
            {
                var paths = pathEnv.Split(Path.PathSeparator);
                foreach (var path in paths)
                {
                    var ffmpegPath = Path.Combine(path, "ffmpeg.exe");
                    if (File.Exists(ffmpegPath))
                    {
                        return ffmpegPath;
                    }
                }
            }

            // 检查常见路径
            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    return path;
                }
            }

            return null;
        }
    }
}




