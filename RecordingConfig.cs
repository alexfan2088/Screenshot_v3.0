using System;
using System.IO;
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

        /// <summary>
        /// 获取视频码率（Mbps）
        /// </summary>
        public int GetVideoBitrateMbps()
        {
            return VideoBitrate switch
            {
                "Low" => 1,
                "Medium" => 3,
                "High" => 5,
                "Auto" => GetAutoVideoBitrate(),
                _ => 3
            };
        }

        /// <summary>
        /// 根据分辨率自动计算码率
        /// </summary>
        private int GetAutoVideoBitrate()
        {
            // 假设原始分辨率为 1920x1080
            int scaledWidth = (int)(1920 * VideoResolutionScale / 100.0);
            int scaledHeight = (int)(1080 * VideoResolutionScale / 100.0);
            int pixels = scaledWidth * scaledHeight;

            // 根据像素数估算码率
            if (pixels < 500000) return 1;      // < 720p
            if (pixels < 1000000) return 2;     // 720p
            return 3;                            // >= 1080p
        }

        /// <summary>
        /// 计算预计文件大小（MB/分钟）
        /// </summary>
        public double EstimateFileSizePerMinute()
        {
            int videoBitrateMbps = GetVideoBitrateMbps();
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
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var config = JsonConvert.DeserializeObject<RecordingConfig>(json);
                    return config ?? new RecordingConfig();
                }
            }
            catch (Exception ex)
            {
                WriteError($"加载配置失败", ex);
            }
            return new RecordingConfig();
        }
    }
}




