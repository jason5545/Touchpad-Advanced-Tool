using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using TouchpadAdvancedTool.Native;
using static TouchpadAdvancedTool.Native.NativeMethods;

namespace TouchpadAdvancedTool.Core
{
    /// <summary>
    /// 滑鼠鉤子管理器 - 負責攔截滑鼠移動事件
    /// </summary>
    public class MouseHookManager : IDisposable
    {
        private readonly ILogger<MouseHookManager> _logger;
        private IntPtr _hookId = IntPtr.Zero;
        private LowLevelMouseProc? _mouseProc;
        private bool _isInstalled;

        /// <summary>
        /// 滑鼠移動事件
        /// </summary>
        public event EventHandler<MouseMoveEventArgs>? MouseMove;

        /// <summary>
        /// 請求攔截滑鼠移動事件（返回 true 表示應該攔截）
        /// </summary>
        public event EventHandler<MouseInterceptEventArgs>? InterceptRequest;

        public MouseHookManager(ILogger<MouseHookManager> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 安裝滑鼠鉤子
        /// </summary>
        public bool InstallHook()
        {
            if (_isInstalled)
            {
                _logger.LogWarning("滑鼠鉤子已安裝");
                return true;
            }

            try
            {
                // 建立鉤子委派（必須保持引用以避免被 GC 回收）
                _mouseProc = HookCallback;

                // 安裝低階滑鼠鉤子
                using (var currentProcess = Process.GetCurrentProcess())
                using (var currentModule = currentProcess.MainModule)
                {
                    if (currentModule?.ModuleName == null)
                    {
                        _logger.LogError("無法取得目前模組名稱");
                        return false;
                    }

                    _hookId = SetWindowsHookEx(
                        HookType.WH_MOUSE_LL,
                        _mouseProc,
                        GetModuleHandle(currentModule.ModuleName),
                        0);
                }

                if (_hookId == IntPtr.Zero)
                {
                    var error = Marshal.GetLastWin32Error();
                    _logger.LogError("安裝滑鼠鉤子失敗，錯誤碼：{ErrorCode}", error);
                    return false;
                }

                _isInstalled = true;
                _logger.LogInformation("滑鼠鉤子已安裝");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "安裝滑鼠鉤子時發生例外");
                return false;
            }
        }

        /// <summary>
        /// 移除滑鼠鉤子
        /// </summary>
        public void UninstallHook()
        {
            if (!_isInstalled || _hookId == IntPtr.Zero)
                return;

            try
            {
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
                _isInstalled = false;
                _logger.LogInformation("滑鼠鉤子已移除");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "移除滑鼠鉤子時發生例外");
            }
        }

        /// <summary>
        /// 鉤子回呼函式
        /// </summary>
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                // nCode >= 0 表示需要處理
                if (nCode >= 0)
                {
                    int message = wParam.ToInt32();

                    // 只處理滑鼠移動事件
                    if (message == WM_MOUSEMOVE)
                    {
                        // 解析滑鼠鉤子結構
                        var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

                        // 觸發滑鼠移動事件
                        var moveArgs = new MouseMoveEventArgs
                        {
                            X = hookStruct.Point.X,
                            Y = hookStruct.Point.Y,
                            Timestamp = DateTime.Now,
                            Flags = hookStruct.Flags,
                            ExtraInfo = hookStruct.ExtraInfo
                        };

                        MouseMove?.Invoke(this, moveArgs);

                        // 詢問是否要攔截此事件
                        var interceptArgs = new MouseInterceptEventArgs
                        {
                            X = hookStruct.Point.X,
                            Y = hookStruct.Point.Y,
                            ShouldIntercept = false
                        };

                        InterceptRequest?.Invoke(this, interceptArgs);

                        // 如果需要攔截，返回非零值阻止事件傳遞
                        if (interceptArgs.ShouldIntercept)
                        {
                            return new IntPtr(1);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "鉤子回呼處理失敗");
            }

            // 呼叫下一個鉤子
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            UninstallHook();
        }
    }

    /// <summary>
    /// 滑鼠移動事件參數
    /// </summary>
    public class MouseMoveEventArgs : EventArgs
    {
        public int X { get; init; }
        public int Y { get; init; }
        public DateTime Timestamp { get; init; }
        public uint Flags { get; init; }
        public IntPtr ExtraInfo { get; init; }
    }

    /// <summary>
    /// 滑鼠攔截事件參數
    /// </summary>
    public class MouseInterceptEventArgs : EventArgs
    {
        public int X { get; init; }
        public int Y { get; init; }
        public bool ShouldIntercept { get; set; }
    }
}
