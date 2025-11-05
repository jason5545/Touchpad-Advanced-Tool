using System;
using System.Runtime.InteropServices;

namespace TouchpadAdvancedTool.Native
{
    /// <summary>
    /// Windows API P/Invoke 宣告
    /// </summary>
    internal static class NativeMethods
    {
        #region Raw Input API

        /// <summary>
        /// 註冊要接收 Raw Input 的裝置
        /// </summary>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool RegisterRawInputDevices(
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] RAWINPUTDEVICE[] pRawInputDevices,
            uint uiNumDevices,
            uint cbSize);

        /// <summary>
        /// 取得 Raw Input 資料
        /// </summary>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetRawInputData(
            IntPtr hRawInput,
            RawInputCommand uiCommand,
            IntPtr pData,
            ref uint pcbSize,
            uint cbSizeHeader);

        /// <summary>
        /// 取得 Raw Input 裝置資訊
        /// </summary>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetRawInputDeviceInfo(
            IntPtr hDevice,
            RawInputDeviceInfo uiCommand,
            IntPtr pData,
            ref uint pcbSize);

        /// <summary>
        /// 列舉所有 Raw Input 裝置
        /// </summary>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetRawInputDeviceList(
            [Out] RAWINPUTDEVICELIST[]? pRawInputDeviceList,
            ref uint puiNumDevices,
            uint cbSize);

        #endregion

        #region HID API

        /// <summary>
        /// 取得 HID 裝置能力
        /// </summary>
        [DllImport("hid.dll", SetLastError = true)]
        public static extern int HidP_GetCaps(
            IntPtr preparsedData,
            ref HIDP_CAPS capabilities);

        /// <summary>
        /// 取得 HID 值能力陣列
        /// </summary>
        [DllImport("hid.dll", SetLastError = true)]
        public static extern int HidP_GetValueCaps(
            HIDP_REPORT_TYPE reportType,
            [Out] HIDP_VALUE_CAPS[] valueCaps,
            ref ushort valueCapsLength,
            IntPtr preparsedData);

        /// <summary>
        /// 取得 HID 按鈕能力陣列
        /// </summary>
        [DllImport("hid.dll", SetLastError = true)]
        public static extern int HidP_GetButtonCaps(
            HIDP_REPORT_TYPE reportType,
            [Out] HIDP_BUTTON_CAPS[] buttonCaps,
            ref ushort buttonCapsLength,
            IntPtr preparsedData);

        /// <summary>
        /// 從 HID 報告中取得使用值
        /// </summary>
        [DllImport("hid.dll", SetLastError = true)]
        public static extern int HidP_GetUsageValue(
            HIDP_REPORT_TYPE reportType,
            ushort usagePage,
            ushort linkCollection,
            ushort usage,
            out uint usageValue,
            IntPtr preparsedData,
            IntPtr report,
            uint reportLength);

        /// <summary>
        /// 從 HID 報告中取得已按下的按鈕
        /// </summary>
        [DllImport("hid.dll", SetLastError = true)]
        public static extern int HidP_GetUsages(
            HIDP_REPORT_TYPE reportType,
            ushort usagePage,
            ushort linkCollection,
            [Out] ushort[] usageList,
            ref uint usageLength,
            IntPtr preparsedData,
            IntPtr report,
            uint reportLength);

        #endregion

        #region Mouse Hook API

        /// <summary>
        /// 安裝 Windows 鉤子
        /// </summary>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(
            HookType idHook,
            LowLevelMouseProc lpfn,
            IntPtr hMod,
            uint dwThreadId);

        /// <summary>
        /// 移除 Windows 鉤子
        /// </summary>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        /// <summary>
        /// 呼叫下一個鉤子
        /// </summary>
        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(
            IntPtr hhk,
            int nCode,
            IntPtr wParam,
            IntPtr lParam);

        /// <summary>
        /// 滑鼠鉤子委派
        /// </summary>
        public delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        #endregion

        #region Input Simulation API

        /// <summary>
        /// 模擬輸入事件
        /// </summary>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(
            uint nInputs,
            [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] INPUT[] pInputs,
            int cbSize);

        #endregion

        #region Helper API

        /// <summary>
        /// 取得模組控制代碼
        /// </summary>
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string? lpModuleName);

        /// <summary>
        /// 取得游標位置
        /// </summary>
        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        /// <summary>
        /// 取得前景視窗
        /// </summary>
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        #endregion

        #region Constants

        // WM 訊息
        public const int WM_INPUT = 0x00FF;
        public const int WM_MOUSEMOVE = 0x0200;
        public const int WM_MOUSEWHEEL = 0x020A;
        public const int WM_MOUSEHWHEEL = 0x020E;
        public const int WM_LBUTTONDOWN = 0x0201;
        public const int WM_LBUTTONUP = 0x0202;
        public const int WM_RBUTTONDOWN = 0x0204;
        public const int WM_RBUTTONUP = 0x0205;

        // Raw Input 裝置類型
        public const uint RIM_TYPEMOUSE = 0;
        public const uint RIM_TYPEKEYBOARD = 1;
        public const uint RIM_TYPEHID = 2;

        // Raw Input 旗標
        public const uint RIDEV_INPUTSINK = 0x00000100;
        public const uint RIDEV_REMOVE = 0x00000001;

        // HID Usage Page
        public const ushort HID_USAGE_PAGE_GENERIC = 0x01;
        public const ushort HID_USAGE_PAGE_DIGITIZER = 0x0D;

        // HID Usage (Generic Desktop)
        public const ushort HID_USAGE_GENERIC_X = 0x30;
        public const ushort HID_USAGE_GENERIC_Y = 0x31;

        // HID Usage (Digitizer)
        public const ushort HID_USAGE_DIGITIZER_TOUCH_PAD = 0x05;
        public const ushort HID_USAGE_DIGITIZER_FINGER = 0x22;
        public const ushort HID_USAGE_DIGITIZER_TIP_SWITCH = 0x42;
        public const ushort HID_USAGE_DIGITIZER_CONFIDENCE = 0x47;
        public const ushort HID_USAGE_DIGITIZER_CONTACT_ID = 0x51;
        public const ushort HID_USAGE_DIGITIZER_CONTACT_COUNT = 0x54;
        public const ushort HID_USAGE_DIGITIZER_SCAN_TIME = 0x56;

        // HID 狀態碼
        public const int HIDP_STATUS_SUCCESS = 0x110000;

        // 滑鼠事件旗標
        public const uint MOUSEEVENTF_MOVE = 0x0001;
        public const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
        public const uint MOUSEEVENTF_WHEEL = 0x0800;
        public const uint MOUSEEVENTF_HWHEEL = 0x1000;

        // 滾輪增量
        public const int WHEEL_DELTA = 120;

        // 輸入類型
        public const uint INPUT_MOUSE = 0;
        public const uint INPUT_KEYBOARD = 1;
        public const uint INPUT_HARDWARE = 2;

        #endregion
    }

    #region Enumerations

    /// <summary>
    /// Raw Input 命令
    /// </summary>
    public enum RawInputCommand : uint
    {
        /// <summary>
        /// 取得輸入資料
        /// </summary>
        RID_INPUT = 0x10000003,

        /// <summary>
        /// 取得標頭資料
        /// </summary>
        RID_HEADER = 0x10000005
    }

    /// <summary>
    /// Raw Input 裝置資訊命令
    /// </summary>
    public enum RawInputDeviceInfo : uint
    {
        /// <summary>
        /// 裝置名稱
        /// </summary>
        RIDI_DEVICENAME = 0x20000007,

        /// <summary>
        /// 裝置資訊
        /// </summary>
        RIDI_DEVICEINFO = 0x2000000b,

        /// <summary>
        /// 預解析資料
        /// </summary>
        RIDI_PREPARSEDDATA = 0x20000005
    }

    /// <summary>
    /// HID 報告類型
    /// </summary>
    public enum HIDP_REPORT_TYPE
    {
        /// <summary>
        /// 輸入報告
        /// </summary>
        HidP_Input = 0,

        /// <summary>
        /// 輸出報告
        /// </summary>
        HidP_Output = 1,

        /// <summary>
        /// 功能報告
        /// </summary>
        HidP_Feature = 2
    }

    /// <summary>
    /// Windows 鉤子類型
    /// </summary>
    public enum HookType : int
    {
        /// <summary>
        /// 低階滑鼠鉤子
        /// </summary>
        WH_MOUSE_LL = 14
    }

    #endregion
}
