using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TouchpadSideScroll.Core;
using TouchpadSideScroll.Models;
using TouchpadSideScroll.ViewModels;

namespace TouchpadSideScroll
{
    /// <summary>
    /// 主視窗
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel? _viewModel;
        private HwndSource? _hwndSource;
        private Hardcodet.Wpf.TaskbarNotification.TaskbarIcon? _notifyIcon;
        private readonly List<Point> _trailPoints = new();
        private const int MaxTrailPoints = 30;
        private Storyboard? _pulseAnimation;

        public MainWindow()
        {
            InitializeComponent();

            // 從資源中取得 NotifyIcon
            _notifyIcon = (Hardcodet.Wpf.TaskbarNotification.TaskbarIcon?)FindResource("NotifyIcon");
        }

        /// <summary>
        /// 視窗載入完成
        /// </summary>
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            try
            {
                // 取得 ViewModel
                _viewModel = App.GetService<MainViewModel>();
                DataContext = _viewModel;

                // 加入視窗訊息鉤子以接收 WM_INPUT
                var windowHandle = new WindowInteropHelper(this).Handle;
                _hwndSource = HwndSource.FromHwnd(windowHandle);
                _hwndSource?.AddHook(WndProc);

                // 訂閱觸控板事件以更新視覺化
                var touchpadTracker = App.GetService<TouchpadTracker>();
                var rawInputManager = App.GetService<RawInputManager>();

                touchpadTracker.EnterScrollZone += OnEnterScrollZone;
                touchpadTracker.ExitScrollZone += OnExitScrollZone;
                rawInputManager.TouchpadInput += OnTouchpadInputForVisualization;
                rawInputManager.TouchpadDetected += OnTouchpadDetectedForVisualization;

                // 設定變更時更新捲動區
                _viewModel.Settings.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(TouchpadSettings.ScrollZoneWidth) ||
                        e.PropertyName == nameof(TouchpadSettings.ScrollZonePosition))
                    {
                        UpdateScrollZoneVisualization();
                    }
                };

                // 初始化 ViewModel
                _viewModel.Initialize(windowHandle);

                // 建立觸控點脈衝動畫
                CreatePulseAnimation();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"視窗初始化失敗：{ex.Message}",
                    "錯誤",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 視窗訊息處理
        /// </summary>
        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_INPUT = 0x00FF;

            if (msg == WM_INPUT)
            {
                _viewModel?.ProcessRawInput(lParam);
            }

            return IntPtr.Zero;
        }

        /// <summary>
        /// 視窗狀態變更
        /// </summary>
        private void Window_StateChanged(object? sender, EventArgs e)
        {
            if (_viewModel?.Settings.MinimizeToTray == true)
            {
                if (WindowState == WindowState.Minimized)
                {
                    // 最小化到系統匣
                    Hide();
                    _notifyIcon?.ShowBalloonTip(
                        "Touchpad 側邊捲動",
                        "應用程式已最小化到系統匣",
                        Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Info);
                }
            }
        }

        /// <summary>
        /// 視窗關閉中
        /// </summary>
        private void Window_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_viewModel?.Settings.MinimizeToTray == true)
            {
                // 取消關閉，改為最小化到系統匣
                e.Cancel = true;
                WindowState = WindowState.Minimized;
            }
        }

        /// <summary>
        /// 系統匣圖示點選
        /// </summary>
        private void NotifyIcon_TrayLeftMouseDown(object sender, RoutedEventArgs e)
        {
            ShowWindow();
        }

        /// <summary>
        /// 顯示視窗
        /// </summary>
        private void ShowWindow_Click(object sender, RoutedEventArgs e)
        {
            ShowWindow();
        }

        /// <summary>
        /// 顯示視窗
        /// </summary>
        private void ShowWindow()
        {
            Show();
            WindowState = WindowState.Normal;
            Activate();
        }

        /// <summary>
        /// 結束應用程式
        /// </summary>
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            // 強制關閉
            if (_viewModel != null)
            {
                _viewModel.Settings.MinimizeToTray = false;
            }
            Application.Current.Shutdown();
        }

        /// <summary>
        /// 超連結導覽
        /// </summary>
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = e.Uri.AbsoluteUri,
                    UseShellExecute = true
                });
                e.Handled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"無法開啟連結：{ex.Message}",
                    "錯誤",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 視窗關閉時清理資源
        /// </summary>
        protected override void OnClosed(EventArgs e)
        {
            _hwndSource?.RemoveHook(WndProc);
            _notifyIcon?.Dispose();
            _pulseAnimation?.Stop();
            base.OnClosed(e);
        }

        #region 觸控板視覺化

        /// <summary>
        /// 建立觸控點脈衝動畫
        /// </summary>
        private void CreatePulseAnimation()
        {
            _pulseAnimation = new Storyboard { RepeatBehavior = RepeatBehavior.Forever };

            var scaleXAnim = new DoubleAnimation
            {
                From = 1.0,
                To = 1.2,
                Duration = TimeSpan.FromSeconds(0.8),
                AutoReverse = true,
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut }
            };
            Storyboard.SetTarget(scaleXAnim, TouchPointScale);
            Storyboard.SetTargetProperty(scaleXAnim, new PropertyPath(ScaleTransform.ScaleXProperty));

            var scaleYAnim = new DoubleAnimation
            {
                From = 1.0,
                To = 1.2,
                Duration = TimeSpan.FromSeconds(0.8),
                AutoReverse = true,
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut }
            };
            Storyboard.SetTarget(scaleYAnim, TouchPointScale);
            Storyboard.SetTargetProperty(scaleYAnim, new PropertyPath(ScaleTransform.ScaleYProperty));

            _pulseAnimation.Children.Add(scaleXAnim);
            _pulseAnimation.Children.Add(scaleYAnim);
        }

        /// <summary>
        /// 觸控板偵測事件處理（用於初始化視覺化）
        /// </summary>
        private void OnTouchpadDetectedForVisualization(object? sender, TouchpadInfo e)
        {
            Dispatcher.Invoke(() =>
            {
                UpdateScrollZoneVisualization();
            });
        }

        /// <summary>
        /// 觸控板輸入事件處理（用於視覺化）
        /// </summary>
        private void OnTouchpadInputForVisualization(object? sender, TouchpadInputEventArgs e)
        {
            if (!_viewModel?.Settings.ShowTouchVisualization ?? true)
                return;

            Dispatcher.Invoke(() =>
            {
                var touchpadInfo = e.TouchpadInfo;
                var contacts = e.Contacts.Where(c => c.IsTouching && c.Confidence).ToList();

                TouchCountText.Text = contacts.Count.ToString();

                if (contacts.Count > 0)
                {
                    var contact = contacts[0]; // 使用第一個觸控點

                    // 計算相對位置 (0-1)
                    double relativeX = (double)(contact.X - touchpadInfo.LogicalMinX) / touchpadInfo.Width;
                    double relativeY = (double)(contact.Y - touchpadInfo.LogicalMinY) / touchpadInfo.Height;

                    // 映射到畫布座標
                    double canvasX = relativeX * TouchpadCanvas.ActualWidth;
                    double canvasY = relativeY * TouchpadCanvas.ActualHeight;

                    // 更新觸控點位置
                    Canvas.SetLeft(TouchPoint, canvasX - TouchPoint.Width / 2);
                    Canvas.SetTop(TouchPoint, canvasY - TouchPoint.Height / 2);

                    if (TouchPoint.Visibility != Visibility.Visible)
                    {
                        TouchPoint.Visibility = Visibility.Visible;
                        HintText.Visibility = Visibility.Collapsed;
                        _pulseAnimation?.Begin();
                    }

                    // 更新位置文字
                    TouchPositionText.Text = $"{relativeX * 100:F1}%, {relativeY * 100:F1}%";

                    // 更新觸控軌跡
                    _trailPoints.Add(new Point(canvasX, canvasY));
                    if (_trailPoints.Count > MaxTrailPoints)
                    {
                        _trailPoints.RemoveAt(0);
                    }

                    TouchTrail.Points = new PointCollection(_trailPoints);
                }
                else
                {
                    // 沒有觸控點時隱藏
                    TouchPoint.Visibility = Visibility.Collapsed;
                    HintText.Visibility = Visibility.Visible;
                    TouchPositionText.Text = "—";
                    _pulseAnimation?.Stop();
                    _trailPoints.Clear();
                    TouchTrail.Points.Clear();
                }
            });
        }

        /// <summary>
        /// 進入捲動區事件處理
        /// </summary>
        private void OnEnterScrollZone(object? sender, ScrollZoneEventArgs e)
        {
            if (!_viewModel?.Settings.ShowTouchVisualization ?? true)
                return;

            Dispatcher.Invoke(() =>
            {
                ScrollStateText.Text = "捲動中";
                ScrollStateText.Foreground = new SolidColorBrush(Color.FromRgb(76, 175, 80)); // 綠色

                // 高亮捲動區
                var fadeIn = new DoubleAnimation
                {
                    To = 0.9,
                    Duration = TimeSpan.FromSeconds(0.2)
                };
                ScrollZoneRect.BeginAnimation(UIElement.OpacityProperty, fadeIn);
            });
        }

        /// <summary>
        /// 離開捲動區事件處理
        /// </summary>
        private void OnExitScrollZone(object? sender, EventArgs e)
        {
            if (!_viewModel?.Settings.ShowTouchVisualization ?? true)
                return;

            Dispatcher.Invoke(() =>
            {
                ScrollStateText.Text = "待機中";
                ScrollStateText.Foreground = SystemColors.ControlTextBrush;

                // 還原捲動區透明度
                var fadeOut = new DoubleAnimation
                {
                    To = 0.6,
                    Duration = TimeSpan.FromSeconds(0.3)
                };
                ScrollZoneRect.BeginAnimation(UIElement.OpacityProperty, fadeOut);
            });
        }

        /// <summary>
        /// 更新捲動區視覺化
        /// </summary>
        private void UpdateScrollZoneVisualization()
        {
            if (_viewModel == null)
                return;

            var rawInputManager = App.GetService<RawInputManager>();
            var touchpadInfo = rawInputManager.ActiveTouchpad;

            if (touchpadInfo == null || !touchpadInfo.IsInitialized)
                return;

            var settings = _viewModel.Settings;
            double zoneWidthPercent = settings.ScrollZoneWidth / 100.0;
            double canvasWidth = TouchpadCanvas.ActualWidth;

            if (canvasWidth == 0)
                canvasWidth = 400; // 預設寬度

            double zoneWidth = canvasWidth * zoneWidthPercent;
            double zoneHeight = TouchpadCanvas.ActualHeight > 0 ? TouchpadCanvas.ActualHeight : 200;

            ScrollZoneRect.Width = zoneWidth;
            ScrollZoneRect.Height = zoneHeight;

            if (settings.ScrollZonePosition == ScrollZonePosition.Right)
            {
                Canvas.SetLeft(ScrollZoneRect, canvasWidth - zoneWidth);
            }
            else
            {
                Canvas.SetLeft(ScrollZoneRect, 0);
            }

            Canvas.SetTop(ScrollZoneRect, 0);
        }

        /// <summary>
        /// 畫布大小變更時更新視覺化
        /// </summary>
        private void TouchpadCanvas_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateScrollZoneVisualization();
        }

        #endregion
    }

    #region Value Converters

    /// <summary>
    /// Bool 轉 Brush 轉換器
    /// </summary>
    public class BoolToBrushConverter : IValueConverter
    {
        public Brush? TrueBrush { get; set; }
        public Brush? FalseBrush { get; set; }

        public object? Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? TrueBrush : FalseBrush;
            }
            return FalseBrush;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Enum 轉 Bool 轉換器（用於 RadioButton）
    /// </summary>
    public class EnumToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return false;

            return value.ToString() == parameter.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue && boolValue && parameter != null)
            {
                return Enum.Parse(targetType, parameter.ToString()!);
            }
            return Binding.DoNothing;
        }
    }

    #endregion
}
