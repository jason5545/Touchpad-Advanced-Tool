using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace TouchpadSideScroll.Models
{
    /// <summary>
    /// 觸控板設定
    /// </summary>
    public class TouchpadSettings : INotifyPropertyChanged
    {
        private bool _isEnabled = true;
        private double _scrollZoneWidth = 15.0;
        private ScrollZonePosition _scrollZonePosition = ScrollZonePosition.Right;
        private double _scrollSensitivity = 1.0;
        private double _scrollSpeed = 2.0;
        private bool _startWithWindows = false;
        private bool _minimizeToTray = true;
        private bool _debugMode = false;
        private bool _invertScrollDirection = false;
        private bool _enableHorizontalScroll = false;
        private bool _invertHorizontalScroll = false;
        private bool _showTouchVisualization = true;

        private int _minimumContactsForScroll = 1;
        private int _maximumContactsForScroll = 1;

        /// <summary>
        /// 是否啟用側邊捲動
        /// </summary>
        public bool IsEnabled
        {
            get => _isEnabled;
            set => SetProperty(ref _isEnabled, value);
        }

        /// <summary>
        /// 捲動區寬度（百分比：5-30）
        /// </summary>
        public double ScrollZoneWidth
        {
            get => _scrollZoneWidth;
            set
            {
                if (value < 5.0) value = 5.0;
                if (value > 30.0) value = 30.0;
                SetProperty(ref _scrollZoneWidth, value);
            }
        }

        /// <summary>
        /// 捲動區位置
        /// </summary>
        public ScrollZonePosition ScrollZonePosition
        {
            get => _scrollZonePosition;
            set => SetProperty(ref _scrollZonePosition, value);
        }

        /// <summary>
        /// 捲動靈敏度（0.1-5.0）
        /// </summary>
        public double ScrollSensitivity
        {
            get => _scrollSensitivity;
            set
            {
                if (value < 0.1) value = 0.1;
                if (value > 5.0) value = 5.0;
                SetProperty(ref _scrollSensitivity, value);
            }
        }

        /// <summary>
        /// 捲動速度（0.5-5.0）
        /// </summary>
        public double ScrollSpeed
        {
            get => _scrollSpeed;
            set
            {
                if (value < 0.5) value = 0.5;
                if (value > 5.0) value = 5.0;
                SetProperty(ref _scrollSpeed, value);
            }
        }

        /// <summary>
        /// 開機自動啟動
        /// </summary>
        public bool StartWithWindows
        {
            get => _startWithWindows;
            set => SetProperty(ref _startWithWindows, value);
        }

        /// <summary>
        /// 最小化到系統匣
        /// </summary>
        public bool MinimizeToTray
        {
            get => _minimizeToTray;
            set => SetProperty(ref _minimizeToTray, value);
        }

        /// <summary>
        /// 除錯模式
        /// </summary>
        public bool DebugMode
        {
            get => _debugMode;
            set => SetProperty(ref _debugMode, value);
        }

        /// <summary>
        /// 反轉捲動方向
        /// </summary>
        public bool InvertScrollDirection
        {
            get => _invertScrollDirection;
            set => SetProperty(ref _invertScrollDirection, value);
        }

        /// <summary>
        /// 反轉水平捲動方向
        /// </summary>
        public bool InvertHorizontalScroll
        {
            get => _invertHorizontalScroll;
            set => SetProperty(ref _invertHorizontalScroll, value);
        }

        /// <summary>
        /// 啟用水平捲動
        /// </summary>
        public bool EnableHorizontalScroll
        {
            get => _enableHorizontalScroll;
            set => SetProperty(ref _enableHorizontalScroll, value);
        }

        /// <summary>
        /// 顯示觸控板視覺化預覽
        /// </summary>
        public bool ShowTouchVisualization
        {
            get => _showTouchVisualization;
            set => SetProperty(ref _showTouchVisualization, value);
        }

        /// <summary>
        /// 最少觸控點數才啟用捲動
        /// </summary>
        public int MinimumContactsForScroll
        {
            get => _minimumContactsForScroll;
            set
            {
                if (value < 1) value = 1;
                if (value > 5) value = 5;
                SetProperty(ref _minimumContactsForScroll, value);
            }
        }

        /// <summary>
        /// 最多觸控點數才啟用捲動
        /// </summary>
        public int MaximumContactsForScroll
        {
            get => _maximumContactsForScroll;
            set
            {
                if (value < 1) value = 1;
                if (value > 5) value = 5;
                SetProperty(ref _maximumContactsForScroll, value);
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
    /// 捲動區位置
    /// </summary>
    public enum ScrollZonePosition
    {
        /// <summary>
        /// 左側
        /// </summary>
        Left,

        /// <summary>
        /// 右側
        /// </summary>
        Right
    }

    /// <summary>
    /// 觸控板資訊
    /// </summary>
    public class TouchpadInfo
    {
        /// <summary>
        /// 裝置控制代碼
        /// </summary>
        public IntPtr DeviceHandle { get; set; }

        /// <summary>
        /// 裝置名稱
        /// </summary>
        public string DeviceName { get; set; } = string.Empty;

        /// <summary>
        /// 廠商 ID
        /// </summary>
        public uint VendorId { get; set; }

        /// <summary>
        /// 產品 ID
        /// </summary>
        public uint ProductId { get; set; }

        /// <summary>
        /// X 軸邏輯最小值
        /// </summary>
        public int LogicalMinX { get; set; }

        /// <summary>
        /// X 軸邏輯最大值
        /// </summary>
        public int LogicalMaxX { get; set; }

        /// <summary>
        /// Y 軸邏輯最小值
        /// </summary>
        public int LogicalMinY { get; set; }

        /// <summary>
        /// Y 軸邏輯最大值
        /// </summary>
        public int LogicalMaxY { get; set; }

        /// <summary>
        /// X 軸物理範圍
        /// </summary>
        public int PhysicalMaxX { get; set; }

        /// <summary>
        /// Y 軸物理範圍
        /// </summary>
        public int PhysicalMaxY { get; set; }

        /// <summary>
        /// 最大觸控點數
        /// </summary>
        public int MaxContactCount { get; set; }

        /// <summary>
        /// 是否為 Precision Touchpad
        /// </summary>
        public bool IsPrecisionTouchpad { get; set; }

        /// <summary>
        /// 觸控板寬度（邏輯單位）
        /// </summary>
        public int Width => LogicalMaxX - LogicalMinX;

        /// <summary>
        /// 觸控板高度（邏輯單位）
        /// </summary>
        public int Height => LogicalMaxY - LogicalMinY;

        /// <summary>
        /// 是否已初始化
        /// </summary>
        public bool IsInitialized => Width > 0 && Height > 0;
    }

    /// <summary>
    /// 觸控點資訊
    /// </summary>
    public class ContactInfo
    {
        /// <summary>
        /// 觸控點 ID
        /// </summary>
        public uint Id { get; set; }

        /// <summary>
        /// X 座標
        /// </summary>
        public int X { get; set; }

        /// <summary>
        /// Y 座標
        /// </summary>
        public int Y { get; set; }

        /// <summary>
        /// 是否按下（Tip Switch）
        /// </summary>
        public bool IsTouching { get; set; }

        /// <summary>
        /// 置信度（用於手掌偵測）
        /// </summary>
        public bool Confidence { get; set; }

        /// <summary>
        /// 時間戳記
        /// </summary>
        public DateTime Timestamp { get; set; }

        /// <summary>
        /// 上一個位置 X
        /// </summary>
        public int LastX { get; set; }

        /// <summary>
        /// 上一個位置 Y
        /// </summary>
        public int LastY { get; set; }

        /// <summary>
        /// X 軸移動距離
        /// </summary>
        public int DeltaX => X - LastX;

        /// <summary>
        /// Y 軸移動距離
        /// </summary>
        public int DeltaY => Y - LastY;
    }
}
