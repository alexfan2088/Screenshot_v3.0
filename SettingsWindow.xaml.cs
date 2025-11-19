using System;
using System.Windows;
using static Screenshot_v3_0.Logger;

namespace Screenshot_v3_0
{
    public partial class SettingsWindow : Window
    {
        private RecordingConfig _config;
        private bool _isInitialized = false; // 标记是否已完成初始化
        public RecordingConfig Config => _config;

        public SettingsWindow(RecordingConfig config)
        {
            // 先设置配置，确保不为空
            _config = config ?? new RecordingConfig();
            
            try
            {
                // 初始化控件
                InitializeComponent();
                
                // 标记初始化完成
                _isInitialized = true;
            }
            catch (Exception ex)
            {
                WriteError($"InitializeComponent 失败", ex);
                WriteLine($"堆栈跟踪: {ex.StackTrace}");
                MessageBox.Show($"初始化控件失败：{ex.Message}\n\n堆栈跟踪：{ex.StackTrace}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                // 即使初始化失败，也尝试继续
                _isInitialized = true; // 标记为已初始化，避免后续错误
            }
            
            try
            {
                // 使用 Loaded 事件确保所有控件都已完全初始化
                this.Loaded += SettingsWindow_Loaded;
            }
            catch (Exception ex)
            {
                WriteError($"注册 Loaded 事件失败", ex);
                // 如果注册事件失败，直接尝试加载配置
                try
                {
                    LoadConfig();
                    UpdateEstimatedSize();
                }
                catch (Exception loadEx)
                {
                    WriteError($"直接加载配置失败", loadEx);
                }
            }
        }

        private void SettingsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // 临时禁用初始化标志，避免在加载配置时触发事件
                bool wasInitialized = _isInitialized;
                _isInitialized = false;
                
                try
                {
                    LoadConfig();
                    UpdateEstimatedSize();
                }
                finally
                {
                    // 恢复初始化标志
                    _isInitialized = wasInitialized;
                }
            }
            catch (Exception ex)
            {
                WriteError($"加载配置失败", ex);
                MessageBox.Show($"加载设置失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadConfig()
        {
            // 加载视频设置 - 分辨率
            try
            {
                if (SliderResolution != null)
                {
                    SliderResolution.Value = _config.VideoResolutionScale;
                }
            }
            catch (Exception ex)
            {
                WriteError($"设置分辨率滑块失败", ex);
            }

            try
            {
                if (TxtResolution != null)
                {
                    TxtResolution.Text = $"{_config.VideoResolutionScale}%";
                }
            }
            catch (Exception ex)
            {
                WriteError($"设置分辨率文本失败", ex);
            }

            // 设置帧率
            try
            {
                if (ComboFrameRate != null && ComboFrameRate.Items != null && ComboFrameRate.Items.Count > 0)
                {
                    foreach (var itemObj in ComboFrameRate.Items)
                    {
                        if (itemObj is System.Windows.Controls.ComboBoxItem item && item?.Tag != null)
                        {
                            if (item.Tag.ToString() == _config.VideoFrameRate.ToString())
                            {
                                ComboFrameRate.SelectedItem = item;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"设置帧率失败", ex);
            }

            // 设置视频码率
            try
            {
                if (ComboVideoBitrate != null && ComboVideoBitrate.Items != null && ComboVideoBitrate.Items.Count > 0)
                {
                    foreach (var itemObj in ComboVideoBitrate.Items)
                    {
                        if (itemObj is System.Windows.Controls.ComboBoxItem item && item?.Tag != null)
                        {
                            if (item.Tag.ToString() == _config.VideoBitrate)
                            {
                                ComboVideoBitrate.SelectedItem = item;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"设置视频码率失败", ex);
            }

            // 设置采样率
            try
            {
                if (ComboSampleRate != null && ComboSampleRate.Items != null && ComboSampleRate.Items.Count > 0)
                {
                    foreach (var itemObj in ComboSampleRate.Items)
                    {
                        if (itemObj is System.Windows.Controls.ComboBoxItem item && item?.Tag != null)
                        {
                            if (item.Tag.ToString() == _config.AudioSampleRate.ToString())
                            {
                                ComboSampleRate.SelectedItem = item;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"设置采样率失败", ex);
            }

            // 设置音频比特率
            try
            {
                if (ComboAudioBitrate != null && ComboAudioBitrate.Items != null && ComboAudioBitrate.Items.Count > 0)
                {
                    foreach (var itemObj in ComboAudioBitrate.Items)
                    {
                        if (itemObj is System.Windows.Controls.ComboBoxItem item && item?.Tag != null)
                        {
                            if (item.Tag.ToString() == _config.AudioBitrate.ToString())
                            {
                                ComboAudioBitrate.SelectedItem = item;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteError($"设置音频比特率失败", ex);
            }
        }

        private void SliderResolution_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // 如果还没有初始化完成，忽略事件
            if (!_isInitialized) return;
            
            try
            {
                int value = (int)e.NewValue;
                if (TxtResolution != null)
                {
                    TxtResolution.Text = $"{value}%";
                }
                if (_config != null)
                {
                    _config.VideoResolutionScale = value;
                }
                UpdateEstimatedSize();
            }
            catch (Exception ex)
            {
                WriteError($"SliderResolution_ValueChanged 失败", ex);
            }
        }

        private void ComboFrameRate_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            // 如果还没有初始化完成，忽略事件
            if (!_isInitialized) return;
            
            try
            {
                if (ComboFrameRate?.SelectedItem is System.Windows.Controls.ComboBoxItem item && item.Tag != null)
                {
                    if (_config != null)
                    {
                        _config.VideoFrameRate = int.Parse(item.Tag.ToString()!);
                    }
                    UpdateEstimatedSize();
                }
            }
            catch (Exception ex)
            {
                WriteError($"ComboFrameRate_SelectionChanged 失败", ex);
            }
        }

        private void ComboVideoBitrate_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            // 如果还没有初始化完成，忽略事件
            if (!_isInitialized) return;
            
            try
            {
                if (ComboVideoBitrate?.SelectedItem is System.Windows.Controls.ComboBoxItem item && item.Tag != null)
                {
                    if (_config != null)
                    {
                        _config.VideoBitrate = item.Tag.ToString()!;
                    }
                    UpdateEstimatedSize();
                }
            }
            catch (Exception ex)
            {
                WriteError($"ComboVideoBitrate_SelectionChanged 失败", ex);
            }
        }

        private void ComboSampleRate_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            // 如果还没有初始化完成，忽略事件
            if (!_isInitialized) return;
            
            try
            {
                if (ComboSampleRate?.SelectedItem is System.Windows.Controls.ComboBoxItem item && item.Tag != null)
                {
                    if (_config != null)
                    {
                        _config.AudioSampleRate = int.Parse(item.Tag.ToString()!);
                    }
                    UpdateEstimatedSize();
                }
            }
            catch (Exception ex)
            {
                WriteError($"ComboSampleRate_SelectionChanged 失败", ex);
            }
        }

        private void ComboAudioBitrate_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            // 如果还没有初始化完成，忽略事件
            if (!_isInitialized) return;
            
            try
            {
                if (ComboAudioBitrate?.SelectedItem is System.Windows.Controls.ComboBoxItem item && item.Tag != null)
                {
                    if (_config != null)
                    {
                        _config.AudioBitrate = int.Parse(item.Tag.ToString()!);
                    }
                    UpdateEstimatedSize();
                }
            }
            catch (Exception ex)
            {
                WriteError($"ComboAudioBitrate_SelectionChanged 失败", ex);
            }
        }

        private void UpdateEstimatedSize()
        {
            try
            {
                if (TxtEstimatedSize != null && _config != null)
                {
                    double sizePerMinute = _config.EstimateFileSizePerMinute();
                    TxtEstimatedSize.Text = $"约 {sizePerMinute:F1} MB/分钟";
                }
            }
            catch (Exception ex)
            {
                WriteError($"更新预计文件大小失败", ex);
            }
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                WriteError($"确定按钮点击失败", ex);
                MessageBox.Show($"保存设置失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DialogResult = false;
                Close();
            }
            catch (Exception ex)
            {
                WriteError($"取消按钮点击失败", ex);
                // 即使出错也强制关闭
                try { Close(); } catch { }
            }
        }
    }
}

