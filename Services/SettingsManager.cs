using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;
using TouchpadAdvancedTool.Models;

namespace TouchpadAdvancedTool.Services
{
    /// <summary>
    /// 設定管理器 - 負責載入、儲存和管理應用程式設定
    /// </summary>
    public class SettingsManager
    {
        private readonly ILogger<SettingsManager> _logger;
        private readonly string _settingsPath;
        private const string RegistryKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "TouchpadAdvancedTool";

        private TouchpadSettings _settings;

        /// <summary>
        /// 目前設定
        /// </summary>
        public TouchpadSettings Settings => _settings;

        /// <summary>
        /// 設定已變更事件
        /// </summary>
        public event EventHandler<TouchpadSettings>? SettingsChanged;

        public SettingsManager(ILogger<SettingsManager> logger)
        {
            _logger = logger;

            // 設定檔路徑：%LocalAppData%\TouchpadAdvancedTool\settings.json
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                AppName);

            Directory.CreateDirectory(appDataPath);
            _settingsPath = Path.Combine(appDataPath, "settings.json");

            // 載入或建立預設設定
            _settings = LoadSettings();

            // 訂閱設定變更事件
            _settings.PropertyChanged += (s, e) =>
            {
                // 自動儲存設定
                SaveSettings();

                // 處理特殊設定
                if (e.PropertyName == nameof(TouchpadSettings.StartWithWindows))
                {
                    UpdateStartupRegistry(_settings.StartWithWindows);
                }

                // 觸發設定變更事件
                SettingsChanged?.Invoke(this, _settings);
            };
        }

        /// <summary>
        /// 載入設定
        /// </summary>
        private TouchpadSettings LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var json = File.ReadAllText(_settingsPath);
                    var options = new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        Converters = { new JsonStringEnumConverter() }
                    };

                    var settings = JsonSerializer.Deserialize<TouchpadSettings>(json, options);

                    if (settings != null)
                    {
                        _logger.LogInformation("已載入設定檔：{Path}", _settingsPath);
                        return settings;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "載入設定檔失敗，使用預設設定");
            }

            _logger.LogInformation("使用預設設定");
            return new TouchpadSettings();
        }

        /// <summary>
        /// 儲存設定
        /// </summary>
        public void SaveSettings()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Converters = { new JsonStringEnumConverter() }
                };

                var json = JsonSerializer.Serialize(_settings, options);
                File.WriteAllText(_settingsPath, json);

                _logger.LogDebug("設定已儲存：{Path}", _settingsPath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "儲存設定檔失敗");
            }
        }

        /// <summary>
        /// 重置為預設設定
        /// </summary>
        public void ResetToDefaults()
        {
            _logger.LogInformation("重置為預設設定");

            _settings.IsEnabled = true;
            _settings.ScrollZoneWidth = 15.0;
            _settings.ScrollZonePosition = ScrollZonePosition.Right;
            _settings.ScrollSensitivity = 1.0;
            _settings.ScrollSpeed = 2.0;
            _settings.InvertScrollDirection = false;
            _settings.EnableHorizontalScroll = false;
            _settings.MinimumContactsForScroll = 1;
            _settings.MaximumContactsForScroll = 1;
            _settings.StartWithWindows = false;
            _settings.MinimizeToTray = true;
            _settings.DebugMode = false;

            SaveSettings();
        }

        /// <summary>
        /// 更新開機自動啟動註冊表項目
        /// </summary>
        private void UpdateStartupRegistry(bool enable)
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, true);

                if (key == null)
                {
                    _logger.LogError("無法開啟註冊表鍵：{Path}", RegistryKeyPath);
                    return;
                }

                if (enable)
                {
                    // 取得應用程式執行檔路徑
                    var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;

                    if (string.IsNullOrEmpty(exePath))
                    {
                        _logger.LogError("無法取得執行檔路徑");
                        return;
                    }

                    // 加入開機啟動項目
                    key.SetValue(AppName, $"\"{exePath}\"", RegistryValueKind.String);
                    _logger.LogInformation("已加入開機啟動項目");
                }
                else
                {
                    // 移除開機啟動項目
                    if (key.GetValue(AppName) != null)
                    {
                        key.DeleteValue(AppName);
                        _logger.LogInformation("已移除開機啟動項目");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "更新開機啟動設定失敗");
            }
        }

        /// <summary>
        /// 檢查是否已設定開機自動啟動
        /// </summary>
        public bool IsStartupEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, false);
                return key?.GetValue(AppName) != null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "檢查開機啟動設定失敗");
                return false;
            }
        }

        /// <summary>
        /// 匯出設定到檔案
        /// </summary>
        public bool ExportSettings(string filePath)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Converters = { new JsonStringEnumConverter() }
                };

                var json = JsonSerializer.Serialize(_settings, options);
                File.WriteAllText(filePath, json);

                _logger.LogInformation("設定已匯出到：{Path}", filePath);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "匯出設定失敗");
                return false;
            }
        }

        /// <summary>
        /// 從檔案匯入設定
        /// </summary>
        public bool ImportSettings(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    _logger.LogWarning("設定檔不存在：{Path}", filePath);
                    return false;
                }

                var json = File.ReadAllText(filePath);
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Converters = { new JsonStringEnumConverter() }
                };

                var settings = JsonSerializer.Deserialize<TouchpadSettings>(json, options);

                if (settings == null)
                {
                    _logger.LogWarning("設定檔格式無效");
                    return false;
                }

                // 複製所有屬性
                _settings.IsEnabled = settings.IsEnabled;
                _settings.ScrollZoneWidth = settings.ScrollZoneWidth;
                _settings.ScrollZonePosition = settings.ScrollZonePosition;
                _settings.ScrollSensitivity = settings.ScrollSensitivity;
                _settings.ScrollSpeed = settings.ScrollSpeed;
                _settings.InvertScrollDirection = settings.InvertScrollDirection;
                _settings.EnableHorizontalScroll = settings.EnableHorizontalScroll;
                _settings.MinimumContactsForScroll = settings.MinimumContactsForScroll;
                _settings.MaximumContactsForScroll = settings.MaximumContactsForScroll;
                _settings.StartWithWindows = settings.StartWithWindows;
                _settings.MinimizeToTray = settings.MinimizeToTray;
                _settings.DebugMode = settings.DebugMode;

                SaveSettings();

                _logger.LogInformation("設定已匯入自：{Path}", filePath);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "匯入設定失敗");
                return false;
            }
        }
    }
}
