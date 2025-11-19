using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Logging;
using TouchpadAdvancedTool.Models;

namespace TouchpadAdvancedTool.Services
{
    /// <summary>
    /// 設定管理器 - 負責載入、儲存和管理應用程式設定
    /// </summary>
    public class SettingsManager
    {
        private readonly ILogger<SettingsManager> _logger;
        private readonly StartupManager _startupManager;
        private readonly string _settingsPath;
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

        public SettingsManager(ILogger<SettingsManager> logger, StartupManager startupManager)
        {
            _logger = logger;
            _startupManager = startupManager;

            // 設定檔路徑：%LocalAppData%\TouchpadAdvancedTool\settings.json
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                AppName);

            Directory.CreateDirectory(appDataPath);
            _settingsPath = Path.Combine(appDataPath, "settings.json");

            // 載入或建立預設設定
            _settings = LoadSettings();

            // 同步開機啟動狀態：檢查實際啟動設定並更新設定
            SyncStartupState();

            // 訂閱設定變更事件
            _settings.PropertyChanged += (s, e) =>
            {
                // 自動儲存設定
                SaveSettings();

                // 處理特殊設定
                if (e.PropertyName == nameof(TouchpadSettings.StartWithWindows))
                {
                    UpdateStartupConfiguration(_settings.StartWithWindows);
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

            // 基本設定
            _settings.IsEnabled = true;
            _settings.ScrollZoneWidth = 15.0;
            _settings.ScrollZonePosition = ScrollZonePosition.Right;
            _settings.ScrollSensitivity = 1.0;
            _settings.ScrollSpeed = 2.0;
            _settings.InvertScrollDirection = false;

            // 水平捲動設定
            _settings.EnableHorizontalScroll = false;
            _settings.InvertHorizontalScroll = false;
            _settings.HorizontalScrollZoneHeight = 15.0;
            _settings.HorizontalScrollZonePosition = HorizontalScrollZonePosition.Top;

            // 觸控點數設定
            _settings.MinimumContactsForScroll = 1;
            _settings.MaximumContactsForScroll = 1;

            // 系統設定
            _settings.StartWithWindows = false;
            _settings.MinimizeToTray = true;
            _settings.DebugMode = false;
            _settings.ShowTouchVisualization = true;

            // 角落觸擊設定
            _settings.EnableCornerTap = false;
            _settings.CornerTapSize = 10.0;
            _settings.CornerTapMaxDuration = 300;
            _settings.CornerTapMovementThreshold = 5.0;
            _settings.TopLeftAction = CornerAction.None;
            _settings.TopRightAction = CornerAction.None;
            _settings.BottomLeftAction = CornerAction.None;
            _settings.BottomRightAction = CornerAction.RightClick;

            SaveSettings();
        }

        /// <summary>
        /// 同步開機啟動狀態
        /// </summary>
        private void SyncStartupState()
        {
            try
            {
                bool startupEnabled = _startupManager.IsStartupEnabled();

                // 如果系統啟動狀態與設定不一致，以系統狀態為準
                if (startupEnabled != _settings.StartWithWindows)
                {
                    _logger.LogInformation("同步開機啟動狀態：System={SystemState}, Settings={SettingsState}, 以系統狀態為準",
                        startupEnabled, _settings.StartWithWindows);

                    // 以系統實際啟動狀態更新設定，避免不一致
                    _settings.StartWithWindows = startupEnabled;
                    SaveSettings();
                }

                // 驗證啟動項目是否有效
                if (startupEnabled && !_startupManager.ValidateStartupEntry())
                {
                    _logger.LogWarning("開機啟動項目無效，正在更新...");
                    _startupManager.EnableStartup(silentStart: true);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "同步開機啟動狀態失敗");
            }
        }

        /// <summary>
        /// 更新開機自動啟動設定（例如 Task Scheduler 排程工作）
        /// </summary>
        private void UpdateStartupConfiguration(bool enable)
        {
            bool success;

            if (enable)
            {
                success = _startupManager.EnableStartup(silentStart: true);
            }
            else
            {
                success = _startupManager.DisableStartup();
            }

            if (!success)
            {
                _logger.LogWarning("更新開機啟動設定失敗");
                System.Windows.MessageBox.Show(
                    "無法更新開機啟動設定，可能需要管理員權限。",
                    "權限不足",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// 檢查是否已設定開機自動啟動
        /// </summary>
        public bool IsStartupEnabled()
        {
            return _startupManager.IsStartupEnabled();
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
                // 基本設定
                _settings.IsEnabled = settings.IsEnabled;
                _settings.ScrollZoneWidth = settings.ScrollZoneWidth;
                _settings.ScrollZonePosition = settings.ScrollZonePosition;
                _settings.ScrollSensitivity = settings.ScrollSensitivity;
                _settings.ScrollSpeed = settings.ScrollSpeed;
                _settings.InvertScrollDirection = settings.InvertScrollDirection;

                // 水平捲動設定
                _settings.EnableHorizontalScroll = settings.EnableHorizontalScroll;
                _settings.InvertHorizontalScroll = settings.InvertHorizontalScroll;
                _settings.HorizontalScrollZoneHeight = settings.HorizontalScrollZoneHeight;
                _settings.HorizontalScrollZonePosition = settings.HorizontalScrollZonePosition;

                // 觸控點數設定
                _settings.MinimumContactsForScroll = settings.MinimumContactsForScroll;
                _settings.MaximumContactsForScroll = settings.MaximumContactsForScroll;

                // 系統設定
                _settings.StartWithWindows = settings.StartWithWindows;
                _settings.MinimizeToTray = settings.MinimizeToTray;
                _settings.DebugMode = settings.DebugMode;
                _settings.ShowTouchVisualization = settings.ShowTouchVisualization;

                // 角落觸擊設定
                _settings.EnableCornerTap = settings.EnableCornerTap;
                _settings.CornerTapSize = settings.CornerTapSize;
                _settings.CornerTapMaxDuration = settings.CornerTapMaxDuration;
                _settings.CornerTapMovementThreshold = settings.CornerTapMovementThreshold;
                _settings.TopLeftAction = settings.TopLeftAction;
                _settings.TopRightAction = settings.TopRightAction;
                _settings.BottomLeftAction = settings.BottomLeftAction;
                _settings.BottomRightAction = settings.BottomRightAction;

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
