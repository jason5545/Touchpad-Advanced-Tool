using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Extensions.Logging;
using TouchpadAdvancedTool.Models;
using TouchpadAdvancedTool.Native;
using static TouchpadAdvancedTool.Native.NativeMethods;

namespace TouchpadAdvancedTool.Core
{
    /// <summary>
    /// 捲動轉換器 - 將觸控板移動轉換為滾輪事件
    /// </summary>
    public class ScrollConverter
    {
        private readonly ILogger<ScrollConverter> _logger;
        private double _accumulatedDeltaY;
        private double _accumulatedDeltaX;
        private DateTime _lastScrollTime;
        private const double MinScrollThreshold = 1.0; // 最小捲動閾值（降低以提升流暢度）

        // 慣性捲動相關
        private readonly Queue<(double deltaY, DateTime time)> _velocityHistory = new(10);
        private Timer? _inertiaTimer;
        private double _currentVelocityY; // 像素/秒
        private double _velocitySumY; // 速度歷史累積值（優化計算）
        private bool _isInertiaScrolling;
        private TouchpadInfo? _lastTouchpadInfo;
        private TouchpadSettings? _lastSettings;

        public ScrollConverter(ILogger<ScrollConverter> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 處理捲動區移動並注入滾輪事件
        /// </summary>
        public void ProcessScroll(ScrollZoneEventArgs args, TouchpadSettings settings)
        {
            try
            {
                // 如果正在慣性捲動，停止它（因為用戶重新開始手動控制）
                if (_isInertiaScrolling)
                {
                    StopInertiaScroll();
                }

                // 儲存觸控板資訊和設定，供慣性捲動使用
                _lastTouchpadInfo = args.TouchpadInfo;
                _lastSettings = settings;

                // 根據捲動區類型決定捲動方向
                int deltaY = settings.InvertScrollDirection ? -args.DeltaY : args.DeltaY;
                int deltaX = settings.InvertHorizontalScroll ? -args.DeltaX : args.DeltaX;

                // 更新速度歷史
                UpdateVelocity(deltaY);

                // 移除未使用的變數
                // int scrollUnitsY = 0;
                // int scrollUnitsX = 0;

                // 計算每個 detent 需要的原始單位數
                double rawPerDetent = ComputeRawUnitsPerDetent(args.TouchpadInfo, settings);
                // 進一步降低閾值比例以提升流暢度（從 0.1 降到 0.05）
                double minThreshold = Math.Max(MinScrollThreshold, rawPerDetent * 0.05);

                if (args.ZoneType == ScrollZoneType.Horizontal)
                {
                    // 水平捲動區：主要使用 X 方向移動來產生水平捲動
                    _accumulatedDeltaX += deltaX;

                    // 優化：預先計算轉換係數
                    double wheelDeltaPerRawUnit = WHEEL_DELTA / rawPerDetent;
                    double potentialWheelDelta = _accumulatedDeltaX * wheelDeltaPerRawUnit;

                    // 最小發送閾值 (例如 15，即 1/8 個 notch)
                    // 許多現代應用程式支援高解析度捲動
                    const double MIN_SEND_DELTA = 15.0;

                    if (Math.Abs(potentialWheelDelta) >= MIN_SEND_DELTA)
                    {
                        int sendAmount = (int)potentialWheelDelta;
                        if (sendAmount != 0)
                        {
                            // 注入滾輪事件 (直接使用計算出的 delta)
                            InjectScrollEvent(0, sendAmount, args.ZoneType, settings, true);
                            
                            // 從累積量中扣除已發送的部分（優化：使用預計算的係數）
                            _accumulatedDeltaX -= sendAmount / wheelDeltaPerRawUnit;

                            if (settings.DebugMode)
                                _logger.LogDebug("水平捲動: 累積={Accum:F2}, 注入Delta={Delta}", _accumulatedDeltaX, sendAmount);
                        }
                    }
                }
                else if (args.ZoneType == ScrollZoneType.Vertical)
                {
                    // 垂直捲動區：主要使用 Y 方向移動來產生垂直捲動
                    _accumulatedDeltaY += deltaY;

                    // 優化：預先計算轉換係數
                    double wheelDeltaPerRawUnit = WHEEL_DELTA / rawPerDetent;
                    double potentialWheelDelta = _accumulatedDeltaY * wheelDeltaPerRawUnit;
                    const double MIN_SEND_DELTA = 15.0;

                    if (Math.Abs(potentialWheelDelta) >= MIN_SEND_DELTA)
                    {
                        int sendAmount = (int)potentialWheelDelta;
                        if (sendAmount != 0)
                        {
                            // 注入滾輪事件
                            InjectScrollEvent(sendAmount, 0, args.ZoneType, settings, true);
                            
                            // 從累積量中扣除已發送的部分（優化：使用預計算的係數）
                            _accumulatedDeltaY -= sendAmount / wheelDeltaPerRawUnit;

                            if (settings.DebugMode)
                                _logger.LogDebug("垂直捲動: 累積={Accum:F2}, 注入Delta={Delta}", _accumulatedDeltaY, sendAmount);
                        }
                    }
                }

                // 注入滾輪事件 (已在上面處理)
                _lastScrollTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "處理捲動失敗");
            }
        }

        /// <summary>
        /// 計算縮放係數
        /// </summary>
        private static double ComputeRawUnitsPerDetent(TouchpadInfo touchpadInfo, TouchpadSettings settings)
        {
            // 標準條件：觸控板高度約 6000 單位；每 40 原始單位 ≈ 1 個 detent (WHEEL_DELTA)
            const double standardHeight = 6000.0;
            const double baseRawUnitsPerDetent = 40.0;

            double height = (touchpadInfo != null && touchpadInfo.IsInitialized) ? touchpadInfo.Height : standardHeight;
            if (height <= 0) height = standardHeight;

            // 依實際高度等比縮放；速度與靈敏度越高，所需原始單位越少
            double rawUnits = baseRawUnitsPerDetent * (height / standardHeight);
            double speed = settings.ScrollSpeed;
            double sensitivity = settings.ScrollSensitivity;
            if (speed < 0.01) speed = 0.01;
            if (sensitivity < 0.01) sensitivity = 0.01;

            rawUnits /= (speed * sensitivity);
            if (rawUnits < 1.0) rawUnits = 1.0;
            return rawUnits;
        }

        /// <summary>
        /// 注入滾輪事件
        /// </summary>
        /// <param name="deltaY">垂直滾輪 Delta (WHEEL_DELTA = 120)</param>
        /// <param name="deltaX">水平滾輪 Delta</param>
        /// <param name="isHighRes">是否為高解析度捲動 (直接使用 delta 值)</param>
        private void InjectScrollEvent(int deltaY, int deltaX, ScrollZoneType zoneType, TouchpadSettings settings, bool isHighRes = false)
        {
            try
            {
                // 垂直捲動
                if (deltaY != 0)
                {
                    var input = new INPUT
                    {
                        Type = INPUT_MOUSE,
                        U = new InputUnion
                        {
                            mi = new MOUSEINPUT
                            {
                                X = 0,
                                Y = 0,
                                MouseData = (uint)(isHighRes ? deltaY : deltaY * WHEEL_DELTA),
                                Flags = MOUSEEVENTF_WHEEL,
                                Time = 0,
                                ExtraInfo = IntPtr.Zero
                            }
                        }
                    };

                    uint result = SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());

                    if (result == 0)
                    {
                        var error = Marshal.GetLastWin32Error();
                        _logger.LogWarning("注入垂直滾輪事件失敗，錯誤碼：{ErrorCode}", error);
                    }
                    else if (settings.DebugMode)
                    {
                        _logger.LogDebug("注入垂直滾輪：{Delta} Delta", deltaY);
                    }
                }

                // 水平捲動
                if (deltaX != 0)
                {
                    var input = new INPUT
                    {
                        Type = INPUT_MOUSE,
                        U = new InputUnion
                        {
                            mi = new MOUSEINPUT
                            {
                                X = 0,
                                Y = 0,
                                MouseData = (uint)(isHighRes ? deltaX : deltaX * WHEEL_DELTA),
                                Flags = MOUSEEVENTF_HWHEEL,
                                Time = 0,
                                ExtraInfo = IntPtr.Zero
                            }
                        }
                    };

                    uint result = SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());

                    if (result == 0)
                    {
                        var error = Marshal.GetLastWin32Error();
                        _logger.LogWarning("注入水平滾輪事件失敗，錯誤碼：{ErrorCode}", error);
                    }
                    else if (settings.DebugMode)
                    {
                        _logger.LogDebug("注入水平滾輪：{Delta} Delta", deltaX);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "注入滾輪事件失敗");
            }
        }

        /// <summary>
        /// 重置累積的捲動距離
        /// </summary>
        public void Reset()
        {
            _accumulatedDeltaY = 0.0;
            _accumulatedDeltaX = 0.0;
            StopInertiaScroll();
        }

        /// <summary>
        /// 更新速度歷史
        /// </summary>
        private void UpdateVelocity(double deltaY)
        {
            var now = DateTime.Now;
            _velocityHistory.Enqueue((deltaY, now));
            _velocitySumY += deltaY; // 優化：維護累積值

            // 只保留最近100ms的歷史
            while (_velocityHistory.Count > 0 && (now - _velocityHistory.Peek().time).TotalMilliseconds > 100)
            {
                var removed = _velocityHistory.Dequeue();
                _velocitySumY -= removed.deltaY; // 優化：從累積值中移除
            }

            // 計算平均速度（像素/秒）
            // 優化：直接使用累積值，無需遍歷 Queue
            if (_velocityHistory.Count >= 2)
            {
                var first = _velocityHistory.First();
                var last = _velocityHistory.Last();
                double totalTime = (last.time - first.time).TotalSeconds;

                if (totalTime > 0)
                {
                    _currentVelocityY = _velocitySumY / totalTime;
                }
            }
        }

        /// <summary>
        /// 開始慣性捲動
        /// </summary>
        public void StartInertiaScroll()
        {
            // 停止現有的慣性捲動
            StopInertiaScroll();

            // 如果速度太小，不啟動慣性（閾值：500像素/秒）
            if (Math.Abs(_currentVelocityY) < 500)
            {
                _logger.LogDebug("慣性速度太小，不啟動慣性捲動：{Velocity:F0}px/s", _currentVelocityY);
                return;
            }

            _isInertiaScrolling = true;
            _logger.LogDebug("開始慣性捲動，初始速度：{Velocity:F0}px/s", _currentVelocityY);

            // 啟動定時器，每16ms（約60fps）執行一次慣性更新
            _inertiaTimer = new Timer(InertiaScrollTick, null, 0, 16);
        }

        /// <summary>
        /// 慣性捲動定時器回調
        /// </summary>
        private void InertiaScrollTick(object? state)
        {
            if (!_isInertiaScrolling) return;

            try
            {
                // 減速係數：每秒減速到原速度的20%（模擬強摩擦力）
                const double decayPerSecond = 0.20;
                double decayPerFrame = Math.Pow(decayPerSecond, 16.0 / 1000.0); // 16ms一幀

                _currentVelocityY *= decayPerFrame;

                // 速度太小時停止（閾值：50像素/秒）
                if (Math.Abs(_currentVelocityY) < 50)
                {
                    _logger.LogDebug("慣性速度降至閾值以下，停止慣性捲動");
                    StopInertiaScroll();
                    return;
                }

                // 計算這一幀應該移動的距離（像素）
                double deltaThisFrame = _currentVelocityY * 0.016; // 16ms = 0.016秒

                // 累積並計算是否注入滾輪事件
                _accumulatedDeltaY += deltaThisFrame;

                // 使用儲存的設定計算 rawPerDetent
                if (_lastTouchpadInfo == null || _lastSettings == null)
                {
                    StopInertiaScroll();
                    return;
                }

                double rawPerDetent = ComputeRawUnitsPerDetent(_lastTouchpadInfo, _lastSettings);
                double minThreshold = Math.Max(MinScrollThreshold, rawPerDetent * 0.05);

                // 計算潛在的滾輪 delta
                double potentialWheelDelta = (_accumulatedDeltaY / rawPerDetent) * WHEEL_DELTA;
                const double MIN_SEND_DELTA = 10.0; // 慣性時使用更低的閾值以保持平滑

                if (Math.Abs(potentialWheelDelta) >= MIN_SEND_DELTA)
                {
                    int sendAmount = (int)potentialWheelDelta;
                    if (sendAmount != 0)
                    {
                        // 注入滾輪事件
                        try
                        {
                            var input = new INPUT
                            {
                                Type = INPUT_MOUSE,
                                U = new InputUnion
                                {
                                    mi = new MOUSEINPUT
                                    {
                                        X = 0,
                                        Y = 0,
                                        MouseData = (uint)sendAmount,
                                        Flags = MOUSEEVENTF_WHEEL,
                                        Time = 0,
                                        ExtraInfo = IntPtr.Zero
                                    }
                                }
                            };

                            SendInput(1, new[] { input }, Marshal.SizeOf<INPUT>());
                            
                            // 從累積量中扣除已發送的部分
                            _accumulatedDeltaY -= sendAmount * (rawPerDetent / WHEEL_DELTA);

                            if (_lastSettings.DebugMode)
                            {
                                _logger.LogDebug("慣性捲動注入：{Delta} Delta，速度：{Velocity:F0}px/s",
                                    sendAmount, _currentVelocityY);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "慣性捲動注入滾輪事件失敗");
                            StopInertiaScroll();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "慣性捲動處理失敗");
                StopInertiaScroll();
            }
        }

        /// <summary>
        /// 停止慣性捲動
        /// </summary>
        public void StopInertiaScroll()
        {
            if (_isInertiaScrolling)
            {
                _logger.LogDebug("停止慣性捲動");
            }

            _isInertiaScrolling = false;
            _inertiaTimer?.Dispose();
            _inertiaTimer = null;
            _velocityHistory.Clear();
            _velocitySumY = 0; // 優化：重置累積值
            _currentVelocityY = 0;
        }
    }
}
