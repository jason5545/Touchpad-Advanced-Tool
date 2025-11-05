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
        /// 目前觸控點數量
        /// </summary>
        public int ActiveContactCount => _activeContacts.Count;

        /// <summary>
        /// 是否正在捲動區內
        /// </summary>
        public bool IsInScrollZone { get; private set; }

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

            // 更新觸控點狀態
            UpdateContacts(args.Contacts);

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
            _primaryContact = _activeContacts.Values.FirstOrDefault(c => c.IsTouching && c.Confidence);

            if (_primaryContact == null)
            {
                if (IsInScrollZone)
                {
                    IsInScrollZone = false;
                    ExitScrollZone?.Invoke(this, EventArgs.Empty);
                }
                return;
            }

            // 檢查是否在捲動區內
            bool inScrollZone = IsContactInScrollZone(_primaryContact, settings);

            if (inScrollZone)
            {
                var scrollArgs = new ScrollZoneEventArgs
                {
                    Contact = _primaryContact,
                    DeltaX = _primaryContact.DeltaX,
                    DeltaY = _primaryContact.DeltaY,
                    TouchpadInfo = _touchpadInfo
                };

                if (!IsInScrollZone)
                {
                    // 進入捲動區
                    IsInScrollZone = true;
                    EnterScrollZone?.Invoke(this, scrollArgs);
                }
                else
                {
                    // 在捲動區內移動
                    ScrollZoneMove?.Invoke(this, scrollArgs);
                }
            }
            else
            {
                if (IsInScrollZone)
                {
                    // 離開捲動區
                    IsInScrollZone = false;
                    ExitScrollZone?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        /// <summary>
        /// 更新觸控點狀態
        /// </summary>
        private void UpdateContacts(List<ContactInfo> contacts)
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
                    contact.LastX = contact.X;
                    contact.LastY = contact.Y;
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
        /// 檢查觸控點是否在捲動區內
        /// </summary>
        private bool IsContactInScrollZone(ContactInfo contact, TouchpadSettings settings)
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
        /// 重置追蹤器狀態
        /// </summary>
        public void Reset()
        {
            _activeContacts.Clear();
            _primaryContact = null;
            IsInScrollZone = false;
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
    }
}
