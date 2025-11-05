using System;
using System.Runtime.InteropServices;

namespace TouchpadAdvancedTool.Native
{
    #region Raw Input Structures

    /// <summary>
    /// Raw Input 裝置註冊結構
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RAWINPUTDEVICE
    {
        /// <summary>
        /// HID Usage Page
        /// </summary>
        public ushort UsagePage;

        /// <summary>
        /// HID Usage
        /// </summary>
        public ushort Usage;

        /// <summary>
        /// 旗標
        /// </summary>
        public uint Flags;

        /// <summary>
        /// 目標視窗控制代碼
        /// </summary>
        public IntPtr Target;
    }

    /// <summary>
    /// Raw Input 資料結構
    /// </summary>
    [StructLayout(LayoutKind.Explicit)]
    public struct RAWINPUT
    {
        /// <summary>
        /// 標頭
        /// </summary>
        [FieldOffset(0)]
        public RAWINPUTHEADER Header;

        /// <summary>
        /// 滑鼠資料
        /// </summary>
        [FieldOffset(24)]
        public RAWMOUSE Mouse;

        /// <summary>
        /// 鍵盤資料
        /// </summary>
        [FieldOffset(24)]
        public RAWKEYBOARD Keyboard;

        /// <summary>
        /// HID 資料
        /// </summary>
        [FieldOffset(24)]
        public RAWHID Hid;
    }

    /// <summary>
    /// Raw Input 標頭
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RAWINPUTHEADER
    {
        /// <summary>
        /// 類型（滑鼠、鍵盤或 HID）
        /// </summary>
        public uint Type;

        /// <summary>
        /// 整個結構大小
        /// </summary>
        public uint Size;

        /// <summary>
        /// 裝置控制代碼
        /// </summary>
        public IntPtr Device;

        /// <summary>
        /// WParam 參數
        /// </summary>
        public IntPtr WParam;
    }

    /// <summary>
    /// Raw Mouse 資料
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RAWMOUSE
    {
        public ushort Flags;
        public uint ButtonFlags;
        public ushort ButtonData;
        public uint RawButtons;
        public int LastX;
        public int LastY;
        public uint ExtraInformation;
    }

    /// <summary>
    /// Raw Keyboard 資料
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RAWKEYBOARD
    {
        public ushort MakeCode;
        public ushort Flags;
        public ushort Reserved;
        public ushort VKey;
        public uint Message;
        public uint ExtraInformation;
    }

    /// <summary>
    /// Raw HID 資料
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RAWHID
    {
        /// <summary>
        /// 單一 HID 報告大小
        /// </summary>
        public uint SizeHid;

        /// <summary>
        /// HID 報告數量
        /// </summary>
        public uint Count;

        /// <summary>
        /// 原始資料（實際上是可變長度陣列）
        /// </summary>
        public byte RawData;
    }

    /// <summary>
    /// Raw Input 裝置列表
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RAWINPUTDEVICELIST
    {
        /// <summary>
        /// 裝置控制代碼
        /// </summary>
        public IntPtr Device;

        /// <summary>
        /// 裝置類型
        /// </summary>
        public uint Type;
    }

    /// <summary>
    /// Raw Input 裝置資訊
    /// </summary>
    [StructLayout(LayoutKind.Explicit)]
    public struct RID_DEVICE_INFO
    {
        [FieldOffset(0)]
        public uint Size;

        [FieldOffset(4)]
        public uint Type;

        [FieldOffset(8)]
        public RID_DEVICE_INFO_MOUSE MouseInfo;

        [FieldOffset(8)]
        public RID_DEVICE_INFO_KEYBOARD KeyboardInfo;

        [FieldOffset(8)]
        public RID_DEVICE_INFO_HID HidInfo;
    }

    /// <summary>
    /// 滑鼠裝置資訊
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RID_DEVICE_INFO_MOUSE
    {
        public uint Id;
        public uint NumberOfButtons;
        public uint SampleRate;
        public bool HasHorizontalWheel;
    }

    /// <summary>
    /// 鍵盤裝置資訊
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RID_DEVICE_INFO_KEYBOARD
    {
        public uint Type;
        public uint SubType;
        public uint KeyboardMode;
        public uint NumberOfFunctionKeys;
        public uint NumberOfIndicators;
        public uint NumberOfKeysTotal;
    }

    /// <summary>
    /// HID 裝置資訊
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct RID_DEVICE_INFO_HID
    {
        public uint VendorId;
        public uint ProductId;
        public uint VersionNumber;
        public ushort UsagePage;
        public ushort Usage;
    }

    #endregion

    #region HID Structures

    /// <summary>
    /// HID 裝置能力
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct HIDP_CAPS
    {
        public ushort Usage;
        public ushort UsagePage;
        public ushort InputReportByteLength;
        public ushort OutputReportByteLength;
        public ushort FeatureReportByteLength;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 17)]
        public ushort[] Reserved;
        public ushort NumberLinkCollectionNodes;
        public ushort NumberInputButtonCaps;
        public ushort NumberInputValueCaps;
        public ushort NumberInputDataIndices;
        public ushort NumberOutputButtonCaps;
        public ushort NumberOutputValueCaps;
        public ushort NumberOutputDataIndices;
        public ushort NumberFeatureButtonCaps;
        public ushort NumberFeatureValueCaps;
        public ushort NumberFeatureDataIndices;
    }

    /// <summary>
    /// HID 值能力
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct HIDP_VALUE_CAPS
    {
        public ushort UsagePage;
        public byte ReportID;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsAlias;
        public ushort BitField;
        public ushort LinkCollection;
        public ushort LinkUsage;
        public ushort LinkUsagePage;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsRange;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsStringRange;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsDesignatorRange;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsAbsolute;
        [MarshalAs(UnmanagedType.U1)]
        public bool HasNull;
        public byte Reserved;
        public ushort BitSize;
        public ushort ReportCount;
        public ushort Reserved2a;
        public ushort Reserved2b;
        public ushort Reserved2c;
        public ushort Reserved2d;
        public ushort Reserved2e;
        public uint UnitsExp;
        public uint Units;
        public int LogicalMin;
        public int LogicalMax;
        public int PhysicalMin;
        public int PhysicalMax;
        public ushort UsageMin;
        public ushort UsageMax;
        public ushort StringMin;
        public ushort StringMax;
        public ushort DesignatorMin;
        public ushort DesignatorMax;
        public ushort DataIndexMin;
        public ushort DataIndexMax;
    }

    /// <summary>
    /// HID 按鈕能力
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct HIDP_BUTTON_CAPS
    {
        public ushort UsagePage;
        public byte ReportID;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsAlias;
        public ushort BitField;
        public ushort LinkCollection;
        public ushort LinkUsage;
        public ushort LinkUsagePage;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsRange;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsStringRange;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsDesignatorRange;
        [MarshalAs(UnmanagedType.U1)]
        public bool IsAbsolute;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 10)]
        public uint[] Reserved;
        public ushort UsageMin;
        public ushort UsageMax;
        public ushort StringMin;
        public ushort StringMax;
        public ushort DesignatorMin;
        public ushort DesignatorMax;
        public ushort DataIndexMin;
        public ushort DataIndexMax;
    }

    #endregion

    #region Mouse Hook Structures

    /// <summary>
    /// 低階滑鼠鉤子結構
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct MSLLHOOKSTRUCT
    {
        /// <summary>
        /// 游標位置
        /// </summary>
        public POINT Point;

        /// <summary>
        /// 滑鼠資料（滾輪時為滾動距離）
        /// </summary>
        public uint MouseData;

        /// <summary>
        /// 事件旗標
        /// </summary>
        public uint Flags;

        /// <summary>
        /// 時間戳記
        /// </summary>
        public uint Time;

        /// <summary>
        /// 額外資訊
        /// </summary>
        public IntPtr ExtraInfo;
    }

    #endregion

    #region Input Simulation Structures

    /// <summary>
    /// 輸入結構
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct INPUT
    {
        /// <summary>
        /// 輸入類型
        /// </summary>
        public uint Type;

        /// <summary>
        /// 輸入聯合（僅使用 Mouse）
        /// </summary>
        public MOUSEINPUT Mouse;
    }

    /// <summary>
    /// 滑鼠輸入
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct MOUSEINPUT
    {
        /// <summary>
        /// X 座標（絕對座標時使用）
        /// </summary>
        public int X;

        /// <summary>
        /// Y 座標（絕對座標時使用）
        /// </summary>
        public int Y;

        /// <summary>
        /// 滑鼠資料（滾輪值或 X 按鈕）
        /// </summary>
        public uint MouseData;

        /// <summary>
        /// 事件旗標
        /// </summary>
        public uint Flags;

        /// <summary>
        /// 時間戳記（0 表示使用系統時間）
        /// </summary>
        public uint Time;

        /// <summary>
        /// 額外資訊
        /// </summary>
        public IntPtr ExtraInfo;
    }

    #endregion

    #region Common Structures

    /// <summary>
    /// 座標點
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        /// <summary>
        /// X 座標
        /// </summary>
        public int X;

        /// <summary>
        /// Y 座標
        /// </summary>
        public int Y;

        public POINT(int x, int y)
        {
            X = x;
            Y = y;
        }
    }

    #endregion
}
