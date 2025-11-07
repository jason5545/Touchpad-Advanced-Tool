using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using Microsoft.Extensions.Logging;
using TouchpadAdvancedTool.Core;
using TouchpadAdvancedTool.Models;
using TouchpadAdvancedTool.Services;

namespace TouchpadAdvancedTool.ViewModels
{
    /// <summary>
    /// 主視窗 ViewModel
    /// </summary>
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly ILogger<MainViewModel> _logger;
        private readonly SettingsManager _settingsManager;
        private readonly RawInputManager _rawInputManager;
        private readonly MouseHookManager _mouseHookManager;
        private readonly TouchpadTracker _touchpadTracker;
        private readonly ScrollConverter _scrollConverter;
        private readonly CornerActionExecutor _cornerActionExecutor;

        private string _statusText = "初始化中...";
        private string _touchpadInfoText = "未偵測到觸控板";
        private string _debugInfoText = string.Empty;
        private bool _isRunning;

        /// <summary>
        /// 設定
        /// </summary>
        public TouchpadSettings Settings { get; }

        /// <summary>
        /// 狀態文字
        /// </summary>
        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }

        /// <summary>
        /// 觸控板資訊文字
        /// </summary>
        public string TouchpadInfoText
        {
            get => _touchpadInfoText;
            set => SetProperty(ref _touchpadInfoText, value);
        }

        /// <summary>
        /// 除錯資訊文字
        /// </summary>
        public string DebugInfoText
        {
            get => _debugInfoText;
            set => SetProperty(ref _debugInfoText, value);
        }

        /// <summary>
        /// 是否正在執行
        /// </summary>
        public bool IsRunning
        {
            get => _isRunning;
            private set => SetProperty(ref _isRunning, value);
        }

        /// <summary>
        /// 切換啟用命令
        /// </summary>
        public ICommand ToggleEnabledCommand { get; }

        /// <summary>
        /// 重置設定命令
        /// </summary>
        public ICommand ResetSettingsCommand { get; }

        /// <summary>
        /// 關閉應用程式命令
        /// </summary>
        public ICommand ExitCommand { get; }

        public MainViewModel(
            ILogger<MainViewModel> logger,
            SettingsManager settingsManager,
            RawInputManager rawInputManager,
            MouseHookManager mouseHookManager,
            TouchpadTracker touchpadTracker,
            ScrollConverter scrollConverter)
        {
            _logger = logger;
            _settingsManager = settingsManager;
            _rawInputManager = rawInputManager;
            _mouseHookManager = mouseHookManager;
            _touchpadTracker = touchpadTracker;
            _scrollConverter = scrollConverter;
            _cornerActionExecutor = new CornerActionExecutor();

            Settings = settingsManager.Settings;

            // 建立命令
            ToggleEnabledCommand = new RelayCommand(ToggleEnabled);
            ResetSettingsCommand = new RelayCommand(ResetSettings);
            ExitCommand = new RelayCommand(Exit);

            // 訂閱事件
            _rawInputManager.TouchpadDetected += OnTouchpadDetected;
            _rawInputManager.TouchpadInput += OnTouchpadInput;
            _touchpadTracker.EnterScrollZone += OnEnterScrollZone;
            _touchpadTracker.ExitScrollZone += OnExitScrollZone;
            _touchpadTracker.ScrollZoneMove += OnScrollZoneMove;
            _touchpadTracker.CornerTap += OnCornerTap;
            _mouseHookManager.InterceptRequest += OnInterceptRequest;

            Settings.PropertyChanged += OnSettingsPropertyChanged;
        }

        /// <summary>
        /// 初始化
        /// </summary>
        public void Initialize(IntPtr windowHandle)
        {
            try
            {
                // 初始化 Raw Input
                bool rawInputOk = _rawInputManager.Initialize(windowHandle);

                if (!rawInputOk)
                {
                    StatusText = "錯誤：無法初始化觸控板輸入";
                    _logger.LogError("Raw Input 初始化失敗");
                    MessageBox.Show(
                        "無法偵測到 Precision Touchpad。\n\n" +
                        "請確認您的觸控板支援 Precision Touchpad 標準，並已安裝 Microsoft 驅動程式。",
                        "TouchpadSideScroll",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                // 安裝滑鼠鉤子
                bool hookOk = _mouseHookManager.InstallHook();

                if (!hookOk)
                {
                    StatusText = "錯誤：無法安裝滑鼠鉤子";
                    _logger.LogError("滑鼠鉤子安裝失敗");
                    MessageBox.Show(
                        "無法安裝滑鼠鉤子。請確認應用程式以管理員權限執行。",
                        "TouchpadSideScroll",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                IsRunning = true;
                UpdateStatus();

                _logger.LogInformation("應用程式初始化完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "初始化失敗");
                StatusText = $"錯誤：{ex.Message}";
                MessageBox.Show(
                    $"初始化失敗：{ex.Message}",
                    "錯誤",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 切換啟用狀態
        /// </summary>
        private void ToggleEnabled()
        {
            Settings.IsEnabled = !Settings.IsEnabled;
        }

        /// <summary>
        /// 重置設定
        /// </summary>
        private void ResetSettings()
        {
            var result = MessageBox.Show(
                "確定要重置所有設定為預設值嗎？",
                "重置設定",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _settingsManager.ResetToDefaults();
                _logger.LogInformation("設定已重置為預設值");
            }
        }

        /// <summary>
        /// 結束應用程式
        /// </summary>
        private void Exit()
        {
            Application.Current.Shutdown();
        }

        /// <summary>
        /// 處理 WM_INPUT 訊息
        /// </summary>
        public void ProcessRawInput(IntPtr lParam)
        {
            _rawInputManager.ProcessInput(lParam);
        }

        /// <summary>
        /// 觸控板偵測事件
        /// </summary>
        private void OnTouchpadDetected(object? sender, TouchpadInfo touchpadInfo)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                // 根據 VID 識別廠商
                string vendorName = GetVendorName(touchpadInfo.VendorId);
                string deviceName = $"{vendorName} Precision Touchpad";

                TouchpadInfoText = $"{deviceName}\n" +
                                   $"範圍：{touchpadInfo.LogicalMinX}~{touchpadInfo.LogicalMaxX} × " +
                                   $"{touchpadInfo.LogicalMinY}~{touchpadInfo.LogicalMaxY}\n" +
                                   $"大小：{touchpadInfo.Width} × {touchpadInfo.Height}";

                UpdateStatus();
            });
        }

        /// <summary>
        /// 根據 VID 取得廠商名稱
        /// </summary>
        private string GetVendorName(uint vid)
        {
            return vid switch
            {
                0x04F3 => "ELAN",           // 義隆電子
                0x06CB => "Synaptics",      // 新思科技
                0x044E => "Alps",           // Alps Electric
                0x2386 => "Raydium",        // 瑞鼎科技
                0x093A => "Pixart",         // 原相科技
                0x1FD2 => "Melfas",         // Melfas
                0x0483 => "STMicro",        // 意法半導體
                0x045E => "Microsoft",      // 微軟
                0x0911 => "Goodix",         // 匯頂科技
                0x27C6 => "Goodix",         // 匯頂科技 (另一個 VID)
                0x048D => "ITE Tech",       // 聯陽半導體
                _ => $"VID:{vid:X4}"        // 未知廠商，顯示 VID
            };
        }

        /// <summary>
        /// 觸控板輸入事件
        /// </summary>
        private void OnTouchpadInput(object? sender, TouchpadInputEventArgs args)
        {
            if (!Settings.IsEnabled)
                return;

            // 更新觸控板追蹤器
            _touchpadTracker.UpdateTouchpadInput(args, Settings);

            // 更新除錯資訊
            if (Settings.DebugMode)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    DebugInfoText = _touchpadTracker.GetDebugInfo();
                });
            }
        }

        /// <summary>
        /// 進入捲動區事件
        /// </summary>
        private void OnEnterScrollZone(object? sender, ScrollZoneEventArgs args)
        {
            _logger.LogDebug("進入捲動區");
            _scrollConverter.Reset();
        }

        /// <summary>
        /// 離開捲動區事件
        /// </summary>
        private void OnExitScrollZone(object? sender, EventArgs args)
        {
            _logger.LogDebug("離開捲動區");
            // 啟動慣性捲動而不是直接重置
            _scrollConverter.StartInertiaScroll();
        }

        /// <summary>
        /// 捲動區移動事件
        /// </summary>
        private void OnScrollZoneMove(object? sender, ScrollZoneEventArgs args)
        {
            if (!Settings.IsEnabled)
                return;

            _scrollConverter.ProcessScroll(args, Settings);
        }

        /// <summary>
        /// 滑鼠攔截請求事件
        /// </summary>
        private void OnInterceptRequest(object? sender, MouseInterceptEventArgs args)
        {
            if (!Settings.IsEnabled)
                return;

            // 快速路徑：如果不在捲動區內，直接返回，不需要進行耗時的來源判斷
            if (!_touchpadTracker.IsInScrollZone)
                return;

            // 只有在捲動區內時，才檢查滑鼠事件是否來自觸控板
            if (_touchpadTracker.IsMouseEventFromTouchpad(Settings))
            {
                args.ShouldIntercept = true;
            }
        }

        /// <summary>
        /// 角落觸擊事件
        /// </summary>
        private void OnCornerTap(object? sender, CornerTapEventArgs args)
        {
            if (!Settings.IsEnabled || !Settings.EnableCornerTap)
                return;

            _logger.LogDebug($"偵測到角落觸擊: {args.Corner}");

            // 根據角落位置取得對應的動作
            CornerAction action = args.Corner switch
            {
                TouchpadCorner.TopLeft => Settings.TopLeftAction,
                TouchpadCorner.TopRight => Settings.TopRightAction,
                TouchpadCorner.BottomLeft => Settings.BottomLeftAction,
                TouchpadCorner.BottomRight => Settings.BottomRightAction,
                _ => CornerAction.None
            };

            if (action != CornerAction.None)
            {
                _logger.LogInformation($"執行角落動作: {args.Corner} -> {action}");
                _cornerActionExecutor.ExecuteAction(action);
            }
        }

        /// <summary>
        /// 設定變更事件
        /// </summary>
        private void OnSettingsPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(TouchpadSettings.IsEnabled))
            {
                UpdateStatus();

                if (!Settings.IsEnabled)
                {
                    _touchpadTracker.Reset();
                    _scrollConverter.Reset();
                }
            }
        }

        /// <summary>
        /// 更新狀態文字
        /// </summary>
        private void UpdateStatus()
        {
            if (!IsRunning)
            {
                StatusText = "未執行";
            }
            else if (Settings.IsEnabled)
            {
                StatusText = "執行中 - 增強功能已啟用";
            }
            else
            {
                StatusText = "執行中 - 增強功能已停用";
            }
        }

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        #endregion
    }

    /// <summary>
    /// 簡單的命令實作
    /// </summary>
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool>? _canExecute;

        public RelayCommand(Action execute, Func<bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object? parameter) => _canExecute?.Invoke() ?? true;

        public void Execute(object? parameter) => _execute();
    }
}
