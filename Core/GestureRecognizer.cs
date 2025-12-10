using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using TouchpadAdvancedTool.Models;

namespace TouchpadAdvancedTool.Core;

/// <summary>
/// 觸控板角落位置列舉
/// </summary>
public enum TouchpadCorner
{
    None,
    TopLeft,
    TopRight,
    BottomLeft,
    BottomRight
}

/// <summary>
/// 角落觸擊事件參數
/// </summary>
public class CornerTapEventArgs : EventArgs
{
    /// <summary>
    /// 觸擊的角落位置
    /// </summary>
    public TouchpadCorner Corner { get; init; }

    /// <summary>
    /// 觸擊時間戳
    /// </summary>
    public DateTime TapTime { get; init; }

    /// <summary>
    /// 觸擊位置 X 座標
    /// </summary>
    public int X { get; init; }

    /// <summary>
    /// 觸擊位置 Y 座標
    /// </summary>
    public int Y { get; init; }
}

/// <summary>
/// 手勢辨識器，負責偵測觸控板上的各種手勢
/// </summary>
public class GestureRecognizer
{
    /// <summary>
    /// 角落觸擊事件
    /// </summary>
    public event EventHandler<CornerTapEventArgs>? CornerTap;

    private readonly TouchpadInfo _touchpadInfo;
    private TouchpadSettings _settings;

    // 觸擊狀態追蹤
    private readonly Dictionary<uint, ContactTapState> _contactTapStates = new();

    /// <summary>
    /// 觸點觸擊狀態
    /// </summary>
    private class ContactTapState
    {
        public TouchpadCorner Corner { get; set; }
        public DateTime TapStartTime { get; set; }
        public int StartX { get; set; }
        public int StartY { get; set; }
        public bool IsValidTap { get; set; }
        public bool CancelledByScroll { get; set; } // 是否因為捲動而取消
    }

    public GestureRecognizer(TouchpadInfo touchpadInfo, TouchpadSettings settings)
    {
        _touchpadInfo = touchpadInfo ?? throw new ArgumentNullException(nameof(touchpadInfo));
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
    }

    /// <summary>
    /// 更新設定
    /// </summary>
    public void UpdateSettings(TouchpadSettings settings)
    {
        _settings = settings ?? throw new ArgumentNullException(nameof(settings));
    }

    /// <summary>
    /// 檢查指定觸點是否在角落區域且正在進行有效的觸擊
    /// </summary>
    public bool IsActiveCornerTap(uint contactId)
    {
        return _contactTapStates.TryGetValue(contactId, out var state) &&
               state.IsValidTap &&
               !state.CancelledByScroll;
    }

    /// <summary>
    /// 檢查是否有任何活動的角落觸擊
    /// </summary>
    public bool HasAnyActiveCornerTap()
    {
        return _contactTapStates.Values.Any(s => s.IsValidTap && !s.CancelledByScroll);
    }

    /// <summary>
    /// 取消指定觸點的角落觸擊（因為開始捲動）
    /// </summary>
    public void CancelCornerTap(uint contactId, string reason = "")
    {
        if (_contactTapStates.TryGetValue(contactId, out var state))
        {
            if (state.IsValidTap && !state.CancelledByScroll)
            {
                state.IsValidTap = false;
                state.CancelledByScroll = true;
                Debug.WriteLine($"[GestureRecognizer] 觸點 {contactId} 的角落觸擊已取消: {reason}");
            }
        }
    }

    /// <summary>
    /// 取得指定觸點所在的角落位置（如果有的話）
    /// </summary>
    public TouchpadCorner GetContactCorner(uint contactId)
    {
        return _contactTapStates.TryGetValue(contactId, out var state) ? state.Corner : TouchpadCorner.None;
    }

    /// <summary>
    /// 處理觸控板輸入，偵測手勢
    /// </summary>
    public void ProcessInput(List<ContactInfo> contacts)
    {
        if (!_settings.EnableCornerTap)
            return;

        foreach (var contact in contacts)
        {
            ProcessContact(contact);
        }

        // 清理已結束的觸點狀態
        CleanupInactiveContacts(contacts);
    }

    /// <summary>
    /// 處理單一觸點
    /// </summary>
    private void ProcessContact(ContactInfo contact)
    {
        if (!contact.Confidence)
            return; // 忽略低置信度觸點（可能是手掌）

        if (contact.IsTouching)
        {
            // 觸控開始或持續中
            if (!_contactTapStates.ContainsKey(contact.Id))
            {
                // 新觸點開始
                var corner = GetTouchpadCorner(contact.X, contact.Y);
                if (corner != TouchpadCorner.None)
                {
                    _contactTapStates[contact.Id] = new ContactTapState
                    {
                        Corner = corner,
                        TapStartTime = contact.Timestamp,
                        StartX = contact.X,
                        StartY = contact.Y,
                        IsValidTap = true,
                        CancelledByScroll = false
                    };

                    Debug.WriteLine($"[GestureRecognizer] 觸點 {contact.Id} 在角落 {corner} 開始觸擊");
                }
            }
            else
            {
                // 觸點持續中，檢查是否移動過多
                var state = _contactTapStates[contact.Id];
                if (state.IsValidTap)
                {
                    if (HasMovedTooMuch(state.StartX, state.StartY, contact.X, contact.Y))
                    {
                        state.IsValidTap = false;
                        Debug.WriteLine($"[GestureRecognizer] 觸點 {contact.Id} 移動過多，取消觸擊");
                    }
                }
            }
        }
        else
        {
            // 觸控結束
            if (_contactTapStates.TryGetValue(contact.Id, out var state))
            {
                if (state.IsValidTap)
                {
                    var duration = (contact.Timestamp - state.TapStartTime).TotalMilliseconds;

                    // 檢查觸擊時長
                    if (duration <= _settings.CornerTapMaxDuration)
                    {
                        // 有效的角落觸擊
                        OnCornerTap(new CornerTapEventArgs
                        {
                            Corner = state.Corner,
                            TapTime = contact.Timestamp,
                            X = contact.X,
                            Y = contact.Y
                        });

                        Debug.WriteLine($"[GestureRecognizer] 觸點 {contact.Id} 完成角落 {state.Corner} 觸擊，時長: {duration:F0}ms");
                    }
                    else
                    {
                        Debug.WriteLine($"[GestureRecognizer] 觸點 {contact.Id} 觸擊時長過長: {duration:F0}ms > {_settings.CornerTapMaxDuration}ms");
                    }
                }

                _contactTapStates.Remove(contact.Id);
            }
        }
    }

    /// <summary>
    /// 判斷觸點是否在角落區域
    /// </summary>
    private TouchpadCorner GetTouchpadCorner(int x, int y)
    {
        int touchpadWidth = _touchpadInfo.Width;
        int touchpadHeight = _touchpadInfo.Height;

        // 計算角落區域大小
        int cornerWidth = (int)(touchpadWidth * _settings.CornerTapSize / 100.0);
        int cornerHeight = (int)(touchpadHeight * _settings.CornerTapSize / 100.0);

        // 判定左右
        bool isLeft = x <= _touchpadInfo.LogicalMinX + cornerWidth;
        bool isRight = x >= _touchpadInfo.LogicalMaxX - cornerWidth;

        // 判定上下
        bool isTop = y <= _touchpadInfo.LogicalMinY + cornerHeight;
        bool isBottom = y >= _touchpadInfo.LogicalMaxY - cornerHeight;

        // 判定角落（必須同時在兩個邊緣，且該角落有設定動作）
        if (isLeft && isTop && _settings.TopLeftAction != CornerAction.None)
            return TouchpadCorner.TopLeft;
        if (isRight && isTop && _settings.TopRightAction != CornerAction.None)
            return TouchpadCorner.TopRight;
        if (isLeft && isBottom && _settings.BottomLeftAction != CornerAction.None)
            return TouchpadCorner.BottomLeft;
        if (isRight && isBottom && _settings.BottomRightAction != CornerAction.None)
            return TouchpadCorner.BottomRight;

        return TouchpadCorner.None;
    }

    /// <summary>
    /// 檢查觸點是否移動過多
    /// </summary>
    private bool HasMovedTooMuch(int startX, int startY, int currentX, int currentY)
    {
        int deltaX = Math.Abs(currentX - startX);
        int deltaY = Math.Abs(currentY - startY);

        // 移動閾值：觸控板寬度/高度的指定百分比
        int thresholdX = (int)(_touchpadInfo.Width * _settings.CornerTapMovementThreshold / 100.0);
        int thresholdY = (int)(_touchpadInfo.Height * _settings.CornerTapMovementThreshold / 100.0);

        return deltaX > thresholdX || deltaY > thresholdY;
    }

    /// <summary>
    /// 清理不活躍的觸點狀態
    /// </summary>
    private void CleanupInactiveContacts(List<ContactInfo> activeContacts)
    {
        var activeIds = new HashSet<uint>();
        foreach (var contact in activeContacts)
        {
            if (contact.IsTouching)
                activeIds.Add(contact.Id);
        }

        // 移除已不存在的觸點狀態
        var idsToRemove = new List<uint>();
        foreach (var id in _contactTapStates.Keys)
        {
            if (!activeIds.Contains(id))
                idsToRemove.Add(id);
        }

        foreach (var id in idsToRemove)
        {
            _contactTapStates.Remove(id);
        }
    }

    /// <summary>
    /// 觸發角落觸擊事件
    /// </summary>
    protected virtual void OnCornerTap(CornerTapEventArgs e)
    {
        CornerTap?.Invoke(this, e);
    }

    /// <summary>
    /// 重設所有手勢狀態
    /// </summary>
    public void Reset()
    {
        _contactTapStates.Clear();
    }
}
