using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Extensions.Logging;
using TouchpadAdvancedTool.Models;

namespace TouchpadAdvancedTool.Core
{
    /// <summary>
    /// 觸控板追蹤器 - 負責追蹤觸控板狀態並判斷觸控位置
    /// </summary>
    public class TouchpadTracker
    {
        private readonly ILogger<TouchpadTracker> _logger;
        private readonly Dictionary<uint, ContactInfo> _activeContacts = new();
        private DateTime _lastTouchpadEventTime;
        private TouchpadInfo? _touchpadInfo;
        private ContactInfo? _primaryContact;
        private GestureRecognizer? _gestureRecognizer;

        // 手勢狀態追蹤
        private enum GestureState
        {
            None,           // 無手勢
            CornerTap,      // 角落觸擊進行中
            Scrolling,      // 捲動進行中
            NormalCursor    // 已離開區域，恢復為一般游標
        }
        private GestureState _currentGestureState = GestureState.None;
        private uint _gestureContactId = 0; // 觸發手勢的觸點 ID

        /// <summary>
        /// 觸控點進入捲動區事件
        /// </summary>
        public event EventHandler<ScrollZoneEventArgs>? EnterScrollZone;

        /// <summary>
        /// 觸控點離開捲動區事件
        /// </summary>
        public event EventHandler? ExitScrollZone;

        /// <summary>
        /// 觸控點在捲動區移動事件
        /// </summary>
        public event EventHandler<ScrollZoneEventArgs>? ScrollZoneMove;

        /// <summary>
        /// 角落觸擊事件
        /// </summary>
        public event EventHandler<CornerTapEventArgs>? CornerTap;

        /// <summary>
        /// 目前觸控點數量
        /// </summary>
        public int ActiveContactCount => _activeContacts.Count;

        /// <summary>
        /// 是否正在捲動區內
        /// </summary>
        public bool IsInScrollZone { get; private set; }

        /// <summary>
        /// 當前捲動區類型
        /// </summary>
        public ScrollZoneType CurrentScrollZoneType { get; private set; } = ScrollZoneType.None;

        /// <summary>
        /// 主要觸控點（用於游標移動）
        /// </summary>
        public ContactInfo? PrimaryContact => _primaryContact;

        public TouchpadTracker(ILogger<TouchpadTracker> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 更新觸控板輸入
        /// </summary>
        public void UpdateTouchpadInput(TouchpadInputEventArgs args, TouchpadSettings settings)
        {
            _lastTouchpadEventTime = DateTime.Now;
            _touchpadInfo = args.TouchpadInfo;

            // 初始化手勢辨識器（如果尚未初始化且啟用角落觸擊）
            if (_gestureRecognizer == null && _touchpadInfo != null && _touchpadInfo.IsInitialized)
            {
                _gestureRecognizer = new GestureRecognizer(_touchpadInfo, settings);
                _gestureRecognizer.CornerTap += OnGestureRecognizerCornerTap;
            }

            // 更新手勢辨識器設定
            _gestureRecognizer?.UpdateSettings(settings);

            // 更新觸控點狀態
            UpdateContacts(args.Contacts, settings);

            // 處理手勢偵測（角落觸擊）
            _gestureRecognizer?.ProcessInput(args.Contacts);

            // 檢查是否所有觸點都已離開（重置手勢狀態）
            if (_activeContacts.Count == 0 || !_activeContacts.Values.Any(c => c.IsTouching))
            {
                if (_currentGestureState != GestureState.None)
                {
                    _logger.LogDebug($"所有觸點離開，重置手勢狀態：{_currentGestureState}");
                    _currentGestureState = GestureState.None;
                    _gestureContactId = 0;
                }

                if (IsInScrollZone)
                {
                    IsInScrollZone = false;
                    CurrentScrollZoneType = ScrollZoneType.None;
                    ExitScrollZone?.Invoke(this, EventArgs.Empty);
                }
                return;
            }

            // 檢查觸控點數量是否符合設定
            if (_activeContacts.Count < settings.MinimumContactsForScroll ||
                _activeContacts.Count > settings.MaximumContactsForScroll)
            {
                // 觸控點數量不符，退出捲動模式
                if (IsInScrollZone)
                {
                    IsInScrollZone = false;
                    ExitScrollZone?.Invoke(this, EventArgs.Empty);
                }
                return;
            }

            // 取得主要觸控點（第一個有效觸控點）
            // 優先選擇高信心的觸控點，但如果沒有，且觸控點在捲動區內，則使用低信心觸控點
            var highConfidenceContact = _activeContacts.Values.FirstOrDefault(c => c.IsTouching && c.Confidence);
            var anyTouchingContact = _activeContacts.Values.FirstOrDefault(c => c.IsTouching);

            // 如果有高信心觸控點，使用它
            if (highConfidenceContact != null)
            {
                _primaryContact = highConfidenceContact;
            }
            // 否則，檢查低信心觸控點是否在捲動區內
            else if (anyTouchingContact != null && GetScrollZoneType(anyTouchingContact, settings) != ScrollZoneType.None)
            {
                _primaryContact = anyTouchingContact;
                if (settings.DebugMode)
                {
                    _logger.LogDebug("使用低信心觸控點（可能是邊緣觸控）：X={X}, Confidence={Confidence}",
                        anyTouchingContact.X, anyTouchingContact.Confidence);
                }
            }
            else
            {
                _primaryContact = null;
            }

            if (_primaryContact == null)
            {
                if (IsInScrollZone)
                {
                    IsInScrollZone = false;
                    ExitScrollZone?.Invoke(this, EventArgs.Empty);
                }
                return;
            }

            // === 手勢狀態管理與優先順序處理 ===

            // 1. 檢查是否在角落區域開始觸控（優先順序最高）
            if (_currentGestureState == GestureState.None &&
                settings.EnableCornerTap &&
                _gestureRecognizer != null &&
                _gestureRecognizer.IsActiveCornerTap(_primaryContact.Id))
            {
                // 在角落區域開始觸控，進入角落觸擊狀態
                _currentGestureState = GestureState.CornerTap;
                _gestureContactId = _primaryContact.Id;
                _logger.LogDebug($"進入角落觸擊狀態，觸點 ID: {_primaryContact.Id}");
                return; // 不處理捲動
            }

            // 2. 如果已經在角落觸擊狀態中
            if (_currentGestureState == GestureState.CornerTap)
            {
                // 檢查是否移動過多，如果是則轉換為捲動狀態或一般游標狀態
                if (_gestureRecognizer != null &&
                    !_gestureRecognizer.IsActiveCornerTap(_gestureContactId))
                {
                    // 角落觸擊已失效（移動過多），檢查是否應該轉換為捲動
                    var scrollZoneType = GetScrollZoneType(_primaryContact, settings);
                    if (scrollZoneType != ScrollZoneType.None)
                    {
                        _logger.LogDebug($"角落觸擊失效，轉換為捲動狀態");
                        _currentGestureState = GestureState.Scrolling;
                        _gestureContactId = _primaryContact.Id;

                        // 取消角落觸擊
                        _gestureRecognizer?.CancelCornerTap(_gestureContactId, "轉換為捲動");

                        // 進入捲動區
                        IsInScrollZone = true;
                        CurrentScrollZoneType = scrollZoneType;
                        EnterScrollZone?.Invoke(this, new ScrollZoneEventArgs
                        {
                            Contact = _primaryContact,
                            DeltaX = _primaryContact.DeltaX,
                            DeltaY = _primaryContact.DeltaY,
                            TouchpadInfo = _touchpadInfo,
                            ZoneType = scrollZoneType
                        });
                    }
                    else
                    {
                        // 不在捲動區，轉換為一般游標狀態
                        _currentGestureState = GestureState.NormalCursor;
                        _gestureContactId = 0;

                        // 確保狀態一致（防止 IsInScrollZone 殘留導致游標被攔截）
                        if (IsInScrollZone)
                        {
                            IsInScrollZone = false;
                            CurrentScrollZoneType = ScrollZoneType.None;
                            ExitScrollZone?.Invoke(this, EventArgs.Empty);
                        }

                        _logger.LogDebug($"離開角落區域，轉換為一般游標模式");
                    }
                }
                return; // 在角落觸擊狀態中，不處理捲動
            }

            // 3. 處理捲動邏輯
            // 快速路徑：如果已經在 NormalCursor 狀態，直接返回，不需要進行區域判斷
            if (_currentGestureState == GestureState.NormalCursor)
            {
                // 確保離開捲動區狀態（防止狀態不一致導致游標被攔截）
                if (IsInScrollZone)
                {
                    IsInScrollZone = false;
                    CurrentScrollZoneType = ScrollZoneType.None;
                    ExitScrollZone?.Invoke(this, EventArgs.Empty);
                }
                return;
            }

            var currentScrollZoneType = GetScrollZoneType(_primaryContact, settings);
            bool inScrollZone = currentScrollZoneType != ScrollZoneType.None;

            if (inScrollZone)
            {

                // 檢查是否在角落區域（如果啟用角落觸擊，則排除角落區域）
                if (settings.EnableCornerTap && IsInCornerZone(_primaryContact, settings))
                {
                    // 在角落區域，不觸發捲動（等待判斷是點擊還是滑動）
                    if (_currentGestureState == GestureState.None)
                    {
                        return;
                    }
                }

                // 進入或保持在捲動狀態
                if (_currentGestureState == GestureState.None)
                {
                    _currentGestureState = GestureState.Scrolling;
                    _gestureContactId = _primaryContact.Id;
                    _logger.LogDebug($"進入捲動狀態，觸點 ID: {_primaryContact.Id}");
                }

                var scrollArgs = new ScrollZoneEventArgs
                {
                    Contact = _primaryContact,
                    DeltaX = _primaryContact.DeltaX,
                    DeltaY = _primaryContact.DeltaY,
                    TouchpadInfo = _touchpadInfo,
                    ZoneType = currentScrollZoneType
                };

                if (!IsInScrollZone || CurrentScrollZoneType != currentScrollZoneType)
                {
                    // 進入捲動區或切換捲動區類型
                    IsInScrollZone = true;
                    CurrentScrollZoneType = currentScrollZoneType;
                    EnterScrollZone?.Invoke(this, scrollArgs);
                }
                else if (_currentGestureState == GestureState.Scrolling)
                {
                    // 在捲動區內移動（只有在 Scrolling 狀態時才觸發）
                    ScrollZoneMove?.Invoke(this, scrollArgs);
                }
            }
            else
            {
                // 離開捲動區
                if (IsInScrollZone)
                {
                    IsInScrollZone = false;
                    CurrentScrollZoneType = ScrollZoneType.None;
                    ExitScrollZone?.Invoke(this, EventArgs.Empty);
                }

                // 在非捲動區時，轉換為一般游標狀態
                // 這確保了從非捲動區進入捲動區時不會觸發捲動（只有從 None 進入才會觸發）
                if (_currentGestureState == GestureState.None ||
                    _currentGestureState == GestureState.Scrolling)
                {
                    _currentGestureState = GestureState.NormalCursor;
                    _logger.LogDebug($"在非捲動區，設置為一般游標模式");
                }
            }
        }

        /// <summary>
        /// 更新觸控點狀態
        /// </summary>
        private void UpdateContacts(List<ContactInfo> contacts, TouchpadSettings settings)
        {
            // 建立當前觸控點 ID 集合
            var currentContactIds = new HashSet<uint>();

            foreach (var contact in contacts)
            {
                currentContactIds.Add(contact.Id);

                if (_activeContacts.TryGetValue(contact.Id, out var existingContact))
                {
                    // 更新現有觸控點的上一個位置
                    contact.LastX = existingContact.X;
                    contact.LastY = existingContact.Y;
                }
                else
                {
                    // 新觸控點，設定初始位置
                    // 為了支援從側邊直接開始滑動，給新觸控點一個初始偏移
                    // 這樣第一幀就能產生有效的 Delta，立即觸發捲動和視覺化
                    contact.LastX = contact.X;
                    contact.LastY = contact.Y;

                    // 如果觸控板已初始化，給予智能初始偏移（僅支援右側捲動區）
                    if (_touchpadInfo != null && _touchpadInfo.IsInitialized)
                    {
                        // 計算觸控點在觸控板上的相對位置（0-1）
                        double xPercent = (double)(contact.X - _touchpadInfo.LogicalMinX) / _touchpadInfo.Width;

                        // 給予一個初始偏移，讓第一幀就能產生滾動
                        const int initialDelta = 20;

                        // 右側捲動區：用戶通常從右往左滑（向內滑動）
                        if (xPercent > 0.7) // 在右側
                        {
                            // 假設用戶將向左滑動，所以 LastX 應該在右邊
                            contact.LastX = contact.X + initialDelta;
                        }
                        else
                        {
                            // 在中央或左側，可能是從中央滑過來的
                            contact.LastX = contact.X - initialDelta;
                        }

                        // Y軸給予初始偏移，假設向下滑動（最常見的情況）
                        // 這樣可以立即產生垂直捲動
                        contact.LastY = contact.Y - initialDelta;
                    }
                }

                _activeContacts[contact.Id] = contact;
            }

            // 移除已不存在的觸控點
            var removedContacts = _activeContacts.Keys.Where(id => !currentContactIds.Contains(id)).ToList();
            foreach (var id in removedContacts)
            {
                _activeContacts.Remove(id);
            }
        }

        /// <summary>
        /// 取得觸控點所在的捲動區類型
        /// </summary>
        private ScrollZoneType GetScrollZoneType(ContactInfo contact, TouchpadSettings settings)
        {
            if (_touchpadInfo == null || !_touchpadInfo.IsInitialized)
                return ScrollZoneType.None;

            // 優先檢查水平捲動區（如果啟用）
            if (settings.EnableHorizontalScroll && IsContactInHorizontalScrollZone(contact, settings))
            {
                return ScrollZoneType.Horizontal;
            }

            // 檢查垂直捲動區
            if (IsContactInVerticalScrollZone(contact, settings))
            {
                return ScrollZoneType.Vertical;
            }

            return ScrollZoneType.None;
        }

        /// <summary>
        /// 檢查觸控點是否在垂直捲動區內
        /// </summary>
        private bool IsContactInVerticalScrollZone(ContactInfo contact, TouchpadSettings settings)
        {
            if (_touchpadInfo == null || !_touchpadInfo.IsInitialized)
                return false;

            // 計算捲動區的 X 座標範圍
            int touchpadWidth = _touchpadInfo.Width;
            double scrollZoneWidthPercent = settings.ScrollZoneWidth / 100.0;
            int scrollZoneWidth = (int)(touchpadWidth * scrollZoneWidthPercent);

            int scrollZoneMinX, scrollZoneMaxX;

            if (settings.ScrollZonePosition == ScrollZonePosition.Right)
            {
                // 右側捲動區
                scrollZoneMinX = _touchpadInfo.LogicalMaxX - scrollZoneWidth;
                scrollZoneMaxX = _touchpadInfo.LogicalMaxX;
            }
            else
            {
                // 左側捲動區
                scrollZoneMinX = _touchpadInfo.LogicalMinX;
                scrollZoneMaxX = _touchpadInfo.LogicalMinX + scrollZoneWidth;
            }

            // 檢查 X 座標是否在捲動區內
            return contact.X >= scrollZoneMinX && contact.X <= scrollZoneMaxX;
        }

        /// <summary>
        /// 檢查觸控點是否在水平捲動區內
        /// </summary>
        private bool IsContactInHorizontalScrollZone(ContactInfo contact, TouchpadSettings settings)
        {
            if (_touchpadInfo == null || !_touchpadInfo.IsInitialized)
                return false;

            // 計算捲動區的 Y 座標範圍
            int touchpadHeight = _touchpadInfo.Height;
            double scrollZoneHeightPercent = settings.HorizontalScrollZoneHeight / 100.0;
            int scrollZoneHeight = (int)(touchpadHeight * scrollZoneHeightPercent);

            int scrollZoneMinY, scrollZoneMaxY;

            if (settings.HorizontalScrollZonePosition == HorizontalScrollZonePosition.Bottom)
            {
                // 底部捲動區
                scrollZoneMinY = _touchpadInfo.LogicalMaxY - scrollZoneHeight;
                scrollZoneMaxY = _touchpadInfo.LogicalMaxY;
            }
            else
            {
                // 頂部捲動區
                scrollZoneMinY = _touchpadInfo.LogicalMinY;
                scrollZoneMaxY = _touchpadInfo.LogicalMinY + scrollZoneHeight;
            }

            // 檢查 Y 座標是否在捲動區內
            return contact.Y >= scrollZoneMinY && contact.Y <= scrollZoneMaxY;
        }

        /// <summary>
        /// 檢查觸控點是否在角落區域內
        /// </summary>
        private bool IsInCornerZone(ContactInfo contact, TouchpadSettings settings)
        {
            if (_touchpadInfo == null || !_touchpadInfo.IsInitialized)
                return false;

            int touchpadWidth = _touchpadInfo.Width;
            int touchpadHeight = _touchpadInfo.Height;

            // 計算角落區域大小
            int cornerWidth = (int)(touchpadWidth * settings.CornerTapSize / 100.0);
            int cornerHeight = (int)(touchpadHeight * settings.CornerTapSize / 100.0);

            // 判定左右
            bool isLeft = contact.X <= _touchpadInfo.LogicalMinX + cornerWidth;
            bool isRight = contact.X >= _touchpadInfo.LogicalMaxX - cornerWidth;

            // 判定上下
            bool isTop = contact.Y <= _touchpadInfo.LogicalMinY + cornerHeight;
            bool isBottom = contact.Y >= _touchpadInfo.LogicalMaxY - cornerHeight;

            // 判定角落（必須同時在兩個邊緣）
            return (isLeft || isRight) && (isTop || isBottom);
        }

        /// <summary>
        /// 檢查滑鼠事件是否來自觸控板
        /// </summary>
        /// <param name="settings">觸控板設定（用於除錯模式）</param>
        /// <returns>如果滑鼠事件來自觸控板則返回 true，否則返回 false</returns>
        public bool IsMouseEventFromTouchpad(TouchpadSettings settings)
        {
            // 使用時間關聯法：如果在短時間內有觸控板事件，則認為滑鼠事件來自觸控板
            // 根據驗證報告建議，將時間視窗從 100ms 縮短到 30ms 以減少誤判
            const int correlationWindowMs = 30; // 30ms 關聯視窗（原為 100ms）
            var timeSinceLastTouchpad = DateTime.Now - _lastTouchpadEventTime;

            // 基本時間檢查
            if (timeSinceLastTouchpad.TotalMilliseconds >= correlationWindowMs)
            {
                if (settings.DebugMode)
                {
                    _logger.LogDebug(
                        "裝置判斷：時間差={TimeDiff:F1}ms 超過閾值 {Threshold}ms，判斷為非觸控板事件",
                        timeSinceLastTouchpad.TotalMilliseconds,
                        correlationWindowMs);
                }
                return false;
            }

            // 檢查是否有活動的觸控點
            if (_activeContacts.Count == 0)
            {
                if (settings.DebugMode)
                {
                    _logger.LogDebug("裝置判斷：無活動觸控點，判斷為非觸控板事件");
                }
                return false;
            }

            // 檢查觸控點是否在移動
            // 如果觸控點靜止但滑鼠在移動，可能是真實滑鼠
            if (_primaryContact != null)
            {
                bool touchpadIsMoving = _primaryContact.DeltaX != 0 || _primaryContact.DeltaY != 0;
                if (!touchpadIsMoving)
                {
                    if (settings.DebugMode)
                    {
                        _logger.LogDebug(
                            "裝置判斷：觸控點靜止（Delta=0），判斷為非觸控板事件");
                    }
                    return false;
                }
            }

            // 所有檢查通過，判斷為觸控板事件
            if (settings.DebugMode)
            {
                _logger.LogDebug(
                    "裝置判斷：時間差={TimeDiff:F1}ms, 觸控點數={ContactCount}, 觸控點移動={IsMoving}, 判斷結果=觸控板",
                    timeSinceLastTouchpad.TotalMilliseconds,
                    _activeContacts.Count,
                    _primaryContact != null ? $"({_primaryContact.DeltaX}, {_primaryContact.DeltaY})" : "N/A");
            }

            return true;
        }

        /// <summary>
        /// 取得目前觸控點的位置資訊（除錯用）
        /// </summary>
        public string GetDebugInfo()
        {
            if (_touchpadInfo == null || !_touchpadInfo.IsInitialized)
                return "觸控板未初始化";

            if (_primaryContact == null)
                return "無觸控點";

            double xPercent = (_primaryContact.X - _touchpadInfo.LogicalMinX) * 100.0 / _touchpadInfo.Width;
            double yPercent = (_primaryContact.Y - _touchpadInfo.LogicalMinY) * 100.0 / _touchpadInfo.Height;

            return $"觸控點數：{_activeContacts.Count}，" +
                   $"位置：({_primaryContact.X}, {_primaryContact.Y})，" +
                   $"百分比：({xPercent:F1}%, {yPercent:F1}%)，" +
                   $"捲動區：{(IsInScrollZone ? "是" : "否")}";
        }

        /// <summary>
        /// 處理手勢辨識器的角落觸擊事件
        /// </summary>
        private void OnGestureRecognizerCornerTap(object? sender, CornerTapEventArgs e)
        {
            // 轉發角落觸擊事件
            CornerTap?.Invoke(this, e);
        }

        /// <summary>
        /// 重置追蹤器狀態
        /// </summary>
        public void Reset()
        {
            _activeContacts.Clear();
            _primaryContact = null;
            IsInScrollZone = false;
            _currentGestureState = GestureState.None;
            _gestureContactId = 0;
            _gestureRecognizer?.Reset();
        }
    }

    /// <summary>
    /// 捲動區事件參數
    /// </summary>
    public class ScrollZoneEventArgs : EventArgs
    {
        public required ContactInfo Contact { get; init; }
        public int DeltaX { get; init; }
        public int DeltaY { get; init; }
        public required TouchpadInfo TouchpadInfo { get; init; }
        public ScrollZoneType ZoneType { get; init; } = ScrollZoneType.Vertical;
    }

    /// <summary>
    /// 捲動區類型
    /// </summary>
    public enum ScrollZoneType
    {
        /// <summary>
        /// 無捲動區
        /// </summary>
        None,

        /// <summary>
        /// 垂直捲動區（左側或右側）
        /// </summary>
        Vertical,

        /// <summary>
        /// 水平捲動區（頂部或底部）
        /// </summary>
        Horizontal
    }
}
