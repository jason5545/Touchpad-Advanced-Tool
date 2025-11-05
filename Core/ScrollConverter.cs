using System;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using TouchpadSideScroll.Models;
using TouchpadSideScroll.Native;
using static TouchpadSideScroll.Native.NativeMethods;

namespace TouchpadSideScroll.Core
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
        private const double MinScrollThreshold = 5.0; // 最小捲動閾值（觸控板原始單位）

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
                // 累積移動距離（支援反向設定）
                int deltaY = settings.InvertScrollDirection ? -args.DeltaY : args.DeltaY;
                int deltaX = settings.InvertHorizontalScroll ? -args.DeltaX : args.DeltaX;

                _accumulatedDeltaY += deltaY;
                if (settings.EnableHorizontalScroll)
                {
                    _accumulatedDeltaX += deltaX;
                }

                // 計算每個 detent 需要的原始單位數
                double rawPerDetent = ComputeRawUnitsPerDetent(args.TouchpadInfo, settings);
                // 門檻：取 detent 的 1/4 或固定最小門檻較大值，避免過度敏感
                double minThreshold = Math.Max(MinScrollThreshold, rawPerDetent * 0.25);

                int scrollUnitsY = 0;
                int scrollUnitsX = 0;

                if (Math.Abs(_accumulatedDeltaY) >= minThreshold)
                {
                    scrollUnitsY = (int)(_accumulatedDeltaY / rawPerDetent);
                    if (scrollUnitsY != 0)
                    {
                        _accumulatedDeltaY -= scrollUnitsY * rawPerDetent; // 保留餘數
                        if (settings.DebugMode)
                            _logger.LogDebug("垂直: 累積={Accum:F2}, detentRaw={Detent:F2}, 注入={Units}", _accumulatedDeltaY, rawPerDetent, scrollUnitsY);
                    }
                }

                if (settings.EnableHorizontalScroll && Math.Abs(_accumulatedDeltaX) >= minThreshold)
                {
                    scrollUnitsX = (int)(_accumulatedDeltaX / rawPerDetent);
                    if (scrollUnitsX != 0)
                    {
                        _accumulatedDeltaX -= scrollUnitsX * rawPerDetent;
                        if (settings.DebugMode)
                            _logger.LogDebug("水平: 累積={Accum:F2}, detentRaw={Detent:F2}, 注入={Units}", _accumulatedDeltaX, rawPerDetent, scrollUnitsX);
                    }
                }

                // 注入滾輪事件
                if (scrollUnitsY != 0 || scrollUnitsX != 0)
                {
                    InjectScrollEvent(scrollUnitsY, scrollUnitsX, settings);
                    _lastScrollTime = DateTime.Now;
                }
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
        private void InjectScrollEvent(int scrollUnitsY, int scrollUnitsX, TouchpadSettings settings)
        {
            try
            {
                // 垂直捲動
                if (scrollUnitsY != 0)
                {
                    var input = new INPUT
                    {
                        Type = INPUT_MOUSE,
                        Mouse = new MOUSEINPUT
                        {
                            X = 0,
                            Y = 0,
                            MouseData = (uint)(scrollUnitsY * WHEEL_DELTA),
                            Flags = MOUSEEVENTF_WHEEL,
                            Time = 0,
                            ExtraInfo = IntPtr.Zero
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
                        _logger.LogDebug("注入垂直滾輪：{Units} 單位", scrollUnitsY);
                    }
                }

                // 水平捲動
                if (scrollUnitsX != 0 && settings.EnableHorizontalScroll)
                {
                    var input = new INPUT
                    {
                        Type = INPUT_MOUSE,
                        Mouse = new MOUSEINPUT
                        {
                            X = 0,
                            Y = 0,
                            MouseData = (uint)(scrollUnitsX * WHEEL_DELTA),
                            Flags = MOUSEEVENTF_HWHEEL,
                            Time = 0,
                            ExtraInfo = IntPtr.Zero
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
                        _logger.LogDebug("注入水平滾輪：{Units} 單位", scrollUnitsX);
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
        }
    }
}
