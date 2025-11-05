using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Extensions.Logging;
using TouchpadAdvancedTool.Models;
using TouchpadAdvancedTool.Native;
using static TouchpadAdvancedTool.Native.NativeMethods;

namespace TouchpadAdvancedTool.Core
{
    /// <summary>
    /// Raw Input 管理器 - 負責接收和解析觸控板原始輸入
    /// </summary>
    public class RawInputManager : IDisposable
    {
        private readonly ILogger<RawInputManager> _logger;
        private IntPtr _windowHandle;
        private bool _isRegistered;
        private readonly Dictionary<IntPtr, TouchpadInfo> _touchpadDevices = new();
        private TouchpadInfo? _activeTouchpad;
        private IntPtr _preparsedData = IntPtr.Zero;

        /// <summary>
        /// 觸控板輸入事件
        /// </summary>
        public event EventHandler<TouchpadInputEventArgs>? TouchpadInput;

        /// <summary>
        /// 觸控板偵測事件
        /// </summary>
        public event EventHandler<TouchpadInfo>? TouchpadDetected;

        /// <summary>
        /// 目前使用的觸控板資訊
        /// </summary>
        public TouchpadInfo? ActiveTouchpad => _activeTouchpad;

        public RawInputManager(ILogger<RawInputManager> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 初始化 Raw Input
        /// </summary>
        public bool Initialize(IntPtr windowHandle)
        {
            try
            {
                _windowHandle = windowHandle;

                // 列舉所有觸控板裝置
                EnumerateTouchpadDevices();

                if (_touchpadDevices.Count == 0)
                {
                    _logger.LogWarning("未偵測到 Precision Touchpad 裝置");
                    return false;
                }

                // 註冊接收 Precision Touchpad 的 Raw Input
                var devices = new RAWINPUTDEVICE[1];
                devices[0].UsagePage = HID_USAGE_PAGE_DIGITIZER;
                devices[0].Usage = HID_USAGE_DIGITIZER_TOUCH_PAD;
                devices[0].Flags = RIDEV_INPUTSINK; // 背景也接收
                devices[0].Target = windowHandle;

                _isRegistered = RegisterRawInputDevices(devices, 1, (uint)Marshal.SizeOf(typeof(RAWINPUTDEVICE)));

                if (!_isRegistered)
                {
                    var error = Marshal.GetLastWin32Error();
                    _logger.LogError("註冊 Raw Input 裝置失敗，錯誤碼：{ErrorCode}", error);
                    return false;
                }

                _logger.LogInformation("成功註冊 Raw Input，偵測到 {Count} 個觸控板裝置", _touchpadDevices.Count);
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "初始化 Raw Input 失敗");
                return false;
            }
        }

        /// <summary>
        /// 處理 WM_INPUT 訊息
        /// </summary>
        public void ProcessInput(IntPtr lParam)
        {
            try
            {
                // 取得 Raw Input 資料大小
                uint size = 0;
                GetRawInputData(lParam, RawInputCommand.RID_INPUT, IntPtr.Zero, ref size, (uint)Marshal.SizeOf<RAWINPUTHEADER>());

                if (size == 0)
                    return;

                // 配置記憶體並取得資料
                IntPtr buffer = Marshal.AllocHGlobal((int)size);
                try
                {
                    uint result = GetRawInputData(lParam, RawInputCommand.RID_INPUT, buffer, ref size, (uint)Marshal.SizeOf<RAWINPUTHEADER>());
                    if (result == unchecked((uint)-1))
                    {
                        _logger.LogWarning("取得 Raw Input 資料失敗");
                        return;
                    }

                    // 解析 RAWINPUT 結構
                    var rawInput = Marshal.PtrToStructure<RAWINPUT>(buffer);

                    // 只處理 HID 裝置
                    if (rawInput.Header.Type == RIM_TYPEHID)
                    {
                        ProcessHidInput(rawInput);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "處理 Raw Input 失敗");
            }
        }

        /// <summary>
        /// 處理 HID 輸入
        /// </summary>
        private void ProcessHidInput(RAWINPUT rawInput)
        {
            try
            {
                // 檢查是否為已知的觸控板裝置
                if (!_touchpadDevices.TryGetValue(rawInput.Header.Device, out var touchpadInfo))
                {
                    // 嘗試偵測新裝置
                    if (DetectTouchpad(rawInput.Header.Device))
                    {
                        touchpadInfo = _touchpadDevices[rawInput.Header.Device];
                    }
                    else
                    {
                        return;
                    }
                }

                // 設定為活動觸控板
                if (_activeTouchpad == null || _activeTouchpad.DeviceHandle != touchpadInfo.DeviceHandle)
                {
                    _activeTouchpad = touchpadInfo;
                    TouchpadDetected?.Invoke(this, touchpadInfo);
                }

                // 取得預解析資料
                if (_preparsedData == IntPtr.Zero)
                {
                    uint size = 0;
                    GetRawInputDeviceInfo(touchpadInfo.DeviceHandle, RawInputDeviceInfo.RIDI_PREPARSEDDATA, IntPtr.Zero, ref size);
                    _preparsedData = Marshal.AllocHGlobal((int)size);
                    GetRawInputDeviceInfo(touchpadInfo.DeviceHandle, RawInputDeviceInfo.RIDI_PREPARSEDDATA, _preparsedData, ref size);
                }

                // 解析 HID 報告
                ParseHidReport(rawInput.Hid, touchpadInfo);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "處理 HID 輸入失敗");
            }
        }

        /// <summary>
        /// 解析 HID 報告
        /// </summary>
        private unsafe void ParseHidReport(RAWHID hidData, TouchpadInfo touchpadInfo)
        {
            try
            {
                // 取得 HID 報告資料的指標
                // hidData.RawData 是結構中的第一個位元組，取得其位址
                IntPtr reportPtr;
                unsafe
                {
                    RAWHID* pHidData = &hidData;
                    reportPtr = new IntPtr(&(pHidData->RawData));
                }
                uint reportLength = hidData.SizeHid;

                // 取得觸控點數量
                int status = HidP_GetUsageValue(
                    HIDP_REPORT_TYPE.HidP_Input,
                    HID_USAGE_PAGE_DIGITIZER,
                    0,
                    HID_USAGE_DIGITIZER_CONTACT_COUNT,
                    out uint contactCount,
                    _preparsedData,
                    reportPtr,
                    reportLength);

                if (status != HIDP_STATUS_SUCCESS)
                {
                    // 某些觸控板可能不回報 Contact Count，預設為 1
                    contactCount = 1;
                }

                var contacts = new List<ContactInfo>();

                // 解析每個觸控點
                for (uint i = 0; i < contactCount && i < 10; i++)
                {
                    var contact = ParseContact(reportPtr, reportLength, i, touchpadInfo);
                    if (contact != null)
                    {
                        contacts.Add(contact);
                    }
                }

                // 觸發觸控板輸入事件
                if (contacts.Count > 0)
                {
                    TouchpadInput?.Invoke(this, new TouchpadInputEventArgs
                    {
                        TouchpadInfo = touchpadInfo,
                        Contacts = contacts,
                        Timestamp = DateTime.Now
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "解析 HID 報告失敗");
            }
        }

        /// <summary>
        /// 解析單一觸控點
        /// </summary>
        private ContactInfo? ParseContact(IntPtr reportPtr, uint reportLength, uint contactIndex, TouchpadInfo touchpadInfo)
        {
            try
            {
                var contact = new ContactInfo
                {
                    Id = contactIndex,
                    Timestamp = DateTime.Now
                };

                // 取得 X 座標
                int status = HidP_GetUsageValue(
                    HIDP_REPORT_TYPE.HidP_Input,
                    HID_USAGE_PAGE_GENERIC,
                    0,
                    HID_USAGE_GENERIC_X,
                    out uint x,
                    _preparsedData,
                    reportPtr,
                    reportLength);

                if (status != HIDP_STATUS_SUCCESS)
                    return null;

                // 取得 Y 座標
                status = HidP_GetUsageValue(
                    HIDP_REPORT_TYPE.HidP_Input,
                    HID_USAGE_PAGE_GENERIC,
                    0,
                    HID_USAGE_GENERIC_Y,
                    out uint y,
                    _preparsedData,
                    reportPtr,
                    reportLength);

                if (status != HIDP_STATUS_SUCCESS)
                    return null;

                contact.X = (int)x;
                contact.Y = (int)y;

                // 取得 Tip Switch（觸控狀態）
                ushort[] usageList = new ushort[10];
                uint usageLength = (uint)usageList.Length;
                status = HidP_GetUsages(
                    HIDP_REPORT_TYPE.HidP_Input,
                    HID_USAGE_PAGE_DIGITIZER,
                    0,
                    usageList,
                    ref usageLength,
                    _preparsedData,
                    reportPtr,
                    reportLength);

                if (status == HIDP_STATUS_SUCCESS)
                {
                    contact.IsTouching = usageList.Take((int)usageLength).Contains(HID_USAGE_DIGITIZER_TIP_SWITCH);
                    contact.Confidence = usageList.Take((int)usageLength).Contains(HID_USAGE_DIGITIZER_CONFIDENCE);
                }
                else
                {
                    // 假設有座標就是在觸控
                    contact.IsTouching = true;
                    contact.Confidence = true;
                }

                return contact;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "解析觸控點 {ContactIndex} 失敗", contactIndex);
                return null;
            }
        }

        /// <summary>
        /// 列舉所有觸控板裝置
        /// </summary>
        private void EnumerateTouchpadDevices()
        {
            try
            {
                uint deviceCount = 0;
                uint size = (uint)Marshal.SizeOf<RAWINPUTDEVICELIST>();

                // 取得裝置數量
                GetRawInputDeviceList(null, ref deviceCount, size);

                if (deviceCount == 0)
                    return;

                // 取得裝置列表
                var devices = new RAWINPUTDEVICELIST[deviceCount];
                GetRawInputDeviceList(devices, ref deviceCount, size);

                // 尋找觸控板裝置
                foreach (var device in devices)
                {
                    if (device.Type == RIM_TYPEHID)
                    {
                        DetectTouchpad(device.Device);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "列舉觸控板裝置失敗");
            }
        }

        /// <summary>
        /// 偵測裝置是否為觸控板
        /// </summary>
        private bool DetectTouchpad(IntPtr deviceHandle)
        {
            try
            {
                // 取得裝置資訊
                uint size = (uint)Marshal.SizeOf<RID_DEVICE_INFO>();
                var deviceInfo = new RID_DEVICE_INFO { Size = size };
                IntPtr deviceInfoPtr = Marshal.AllocHGlobal((int)size);

                try
                {
                    Marshal.StructureToPtr(deviceInfo, deviceInfoPtr, false);
                    uint result = GetRawInputDeviceInfo(deviceHandle, RawInputDeviceInfo.RIDI_DEVICEINFO, deviceInfoPtr, ref size);

                    if (result == unchecked((uint)-1))
                        return false;

                    deviceInfo = Marshal.PtrToStructure<RID_DEVICE_INFO>(deviceInfoPtr);

                    // 檢查是否為 Precision Touchpad
                    if (deviceInfo.Type == RIM_TYPEHID &&
                        deviceInfo.HidInfo.UsagePage == HID_USAGE_PAGE_DIGITIZER &&
                        deviceInfo.HidInfo.Usage == HID_USAGE_DIGITIZER_TOUCH_PAD)
                    {
                        var touchpadInfo = new TouchpadInfo
                        {
                            DeviceHandle = deviceHandle,
                            VendorId = deviceInfo.HidInfo.VendorId,
                            ProductId = deviceInfo.HidInfo.ProductId,
                            IsPrecisionTouchpad = true
                        };

                        // 取得裝置名稱
                        touchpadInfo.DeviceName = GetDeviceName(deviceHandle);

                        // 取得座標範圍
                        GetTouchpadCapabilities(deviceHandle, touchpadInfo);

                        _touchpadDevices[deviceHandle] = touchpadInfo;

                        _logger.LogInformation(
                            "偵測到觸控板：{Name}，VID={VendorId:X4}，PID={ProductId:X4}，範圍={MinX}~{MaxX} × {MinY}~{MaxY}",
                            touchpadInfo.DeviceName,
                            touchpadInfo.VendorId,
                            touchpadInfo.ProductId,
                            touchpadInfo.LogicalMinX,
                            touchpadInfo.LogicalMaxX,
                            touchpadInfo.LogicalMinY,
                            touchpadInfo.LogicalMaxY);

                        return true;
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(deviceInfoPtr);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "偵測觸控板失敗");
            }

            return false;
        }

        /// <summary>
        /// 取得裝置名稱
        /// </summary>
        private string GetDeviceName(IntPtr deviceHandle)
        {
            try
            {
                uint size = 0;
                GetRawInputDeviceInfo(deviceHandle, RawInputDeviceInfo.RIDI_DEVICENAME, IntPtr.Zero, ref size);

                if (size == 0)
                    return "Unknown Touchpad";

                IntPtr namePtr = Marshal.AllocHGlobal((int)size * 2); // Unicode
                try
                {
                    GetRawInputDeviceInfo(deviceHandle, RawInputDeviceInfo.RIDI_DEVICENAME, namePtr, ref size);
                    return Marshal.PtrToStringUni(namePtr) ?? "Unknown Touchpad";
                }
                finally
                {
                    Marshal.FreeHGlobal(namePtr);
                }
            }
            catch
            {
                return "Unknown Touchpad";
            }
        }

        /// <summary>
        /// 取得觸控板能力（座標範圍）
        /// </summary>
        private void GetTouchpadCapabilities(IntPtr deviceHandle, TouchpadInfo touchpadInfo)
        {
            try
            {
                // 取得預解析資料
                uint size = 0;
                GetRawInputDeviceInfo(deviceHandle, RawInputDeviceInfo.RIDI_PREPARSEDDATA, IntPtr.Zero, ref size);

                if (size == 0)
                    return;

                IntPtr preparsedData = Marshal.AllocHGlobal((int)size);
                try
                {
                    GetRawInputDeviceInfo(deviceHandle, RawInputDeviceInfo.RIDI_PREPARSEDDATA, preparsedData, ref size);

                    // 取得 HID 能力
                    var caps = new HIDP_CAPS();
                    int status = HidP_GetCaps(preparsedData, ref caps);

                    if (status != HIDP_STATUS_SUCCESS)
                        return;

                    // 取得值能力陣列
                    ushort valueCapsLength = caps.NumberInputValueCaps;
                    if (valueCapsLength == 0)
                        return;

                    var valueCaps = new HIDP_VALUE_CAPS[valueCapsLength];
                    status = HidP_GetValueCaps(HIDP_REPORT_TYPE.HidP_Input, valueCaps, ref valueCapsLength, preparsedData);

                    if (status != HIDP_STATUS_SUCCESS)
                        return;

                    // 尋找 X 和 Y 軸的座標範圍
                    foreach (var cap in valueCaps)
                    {
                        if (cap.UsagePage == HID_USAGE_PAGE_GENERIC)
                        {
                            if (cap.UsageMin == HID_USAGE_GENERIC_X || cap.UsageMax == HID_USAGE_GENERIC_X)
                            {
                                touchpadInfo.LogicalMinX = cap.LogicalMin;
                                touchpadInfo.LogicalMaxX = cap.LogicalMax;
                                touchpadInfo.PhysicalMaxX = cap.PhysicalMax;
                            }
                            else if (cap.UsageMin == HID_USAGE_GENERIC_Y || cap.UsageMax == HID_USAGE_GENERIC_Y)
                            {
                                touchpadInfo.LogicalMinY = cap.LogicalMin;
                                touchpadInfo.LogicalMaxY = cap.LogicalMax;
                                touchpadInfo.PhysicalMaxY = cap.PhysicalMax;
                            }
                        }
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(preparsedData);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "取得觸控板能力失敗");
            }
        }

        /// <summary>
        /// 解除註冊 Raw Input
        /// </summary>
        public void Unregister()
        {
            if (!_isRegistered)
                return;

            try
            {
                var devices = new RAWINPUTDEVICE[1];
                devices[0].UsagePage = HID_USAGE_PAGE_DIGITIZER;
                devices[0].Usage = HID_USAGE_DIGITIZER_TOUCH_PAD;
                devices[0].Flags = RIDEV_REMOVE;
                devices[0].Target = IntPtr.Zero;

                RegisterRawInputDevices(devices, 1, (uint)Marshal.SizeOf(typeof(RAWINPUTDEVICE)));
                _isRegistered = false;

                _logger.LogInformation("已解除註冊 Raw Input");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "解除註冊 Raw Input 失敗");
            }
        }

        public void Dispose()
        {
            Unregister();

            if (_preparsedData != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(_preparsedData);
                _preparsedData = IntPtr.Zero;
            }
        }
    }

    /// <summary>
    /// 觸控板輸入事件參數
    /// </summary>
    public class TouchpadInputEventArgs : EventArgs
    {
        /// <summary>
        /// 觸控板資訊
        /// </summary>
        public required TouchpadInfo TouchpadInfo { get; init; }

        /// <summary>
        /// 觸控點列表
        /// </summary>
        public required List<ContactInfo> Contacts { get; init; }

        /// <summary>
        /// 時間戳記
        /// </summary>
        public DateTime Timestamp { get; init; }
    }
}
