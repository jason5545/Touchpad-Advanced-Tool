# Touchpad Advanced Tool

一個使用 Windows Precision Touchpad API 實作的觸控板進階工具，提供側邊捲動等多種觸控板增強功能，讓您可以在觸控板的邊緣區域垂直滑動來捲動頁面，就像 MacOS 的側邊捲動功能一樣。

## 功能特色

- ✨ **側邊捲動**：在觸控板右側（或左側）區域垂直滑動即可捲動頁面
- 🎯 **Precision Touchpad 支援**：使用 Windows Precision Touchpad API 取得原始觸控資料
- ⚙️ **可自訂設定**：調整捲動區寬度、位置、速度和靈敏度
- 🎨 **現代化 UI**：使用 ModernWpf 打造的美觀介面
- 🔧 **進階選項**：支援開機自動啟動、系統匣最小化、除錯模式等
- 🌙 **主題支援**：自動跟隨系統深色/淺色主題

## 系統需求

- Windows 10 或 Windows 11
- 支援 Precision Touchpad 的觸控板（需安裝 Microsoft 驅動程式）
- .NET 8.0 或更新版本

## 安裝與使用

### 編譯專案

1. 克隆此儲存庫：
   ```bash
   git clone https://github.com/yourusername/Touchpad-sidescroll.git
   cd Touchpad-sidescroll
   ```

2. 使用 Visual Studio 2022 或更新版本開啟 `TouchpadAdvancedTool.csproj`

3. 建置專案（需要 .NET 8.0 SDK）：
   ```bash
   dotnet build -c Release
   ```

4. **以管理員權限執行**（安裝全域滑鼠鉤子需要管理員權限）：
   ```bash
   cd bin/Release/net8.0-windows
   # 以管理員身分執行
   .\TouchpadAdvancedTool.exe
   ```

### 使用說明

1. 首次啟動時，應用程式會自動偵測您的 Precision Touchpad
2. 預設設定為**右側 15% 區域**作為捲動區
3. 在捲動區內**垂直滑動**即可捲動頁面
4. 可在主視窗中調整各項設定：
   - **捲動區寬度**：5% ~ 30%
   - **捲動區位置**：左側或右側
   - **捲動速度**：0.5x ~ 5.0x
   - **捲動靈敏度**：0.1x ~ 5.0x

## 技術實作

### 核心技術

1. **Raw Input API**
   - 使用 `RegisterRawInputDevices` 註冊接收 Precision Touchpad 的原始輸入
   - 解析 HID Report Descriptor 取得觸控板座標範圍
   - 追蹤觸控點的絕對座標位置

2. **滑鼠低階鉤子 (WH_MOUSE_LL)**
   - 攔截滑鼠移動事件
   - 使用時間關聯法判斷滑鼠事件是否來自觸控板
   - 在捲動區內時攔截游標移動

3. **滾輪事件注入**
   - 使用 `SendInput` API 注入 MOUSEEVENTF_WHEEL 事件
   - 將觸控板 Y 軸移動量轉換為滾輪增量
   - 支援垂直和水平捲動

### 專案結構

```
TouchpadAdvancedTool/
├── Native/                     # P/Invoke 宣告和結構定義
│   ├── NativeMethods.cs        # Windows API 宣告
│   └── Structures.cs           # 原生結構定義
├── Core/                       # 核心邏輯
│   ├── RawInputManager.cs      # Raw Input 處理
│   ├── MouseHookManager.cs     # 滑鼠鉤子管理
│   ├── TouchpadTracker.cs      # 觸控板狀態追蹤
│   └── ScrollConverter.cs      # 捲動轉換器
├── Models/                     # 資料模型
│   └── TouchpadSettings.cs     # 設定和觸控板資訊
├── Services/                   # 服務層
│   └── SettingsManager.cs      # 設定管理
├── ViewModels/                 # ViewModel
│   └── MainViewModel.cs        # 主視窗 ViewModel
├── Resources/                  # 資源檔案
│   └── Styles.xaml             # 樣式定義
├── App.xaml / App.xaml.cs      # 應用程式進入點
└── MainWindow.xaml/.cs         # 主視窗
```

## 相容性

### 支援的觸控板

- Microsoft Precision Touchpad
- Synaptics 觸控板（需安裝 Microsoft Precision Touchpad 驅動）
- ELAN 觸控板（需安裝 Microsoft Precision Touchpad 驅動）
- 其他支援 Precision Touchpad 標準的裝置

### 已知限制

1. **需要管理員權限**：安裝全域滑鼠鉤子需要提升權限
2. **Precision Touchpad 限制**：只支援使用 Microsoft Precision Touchpad 驅動的裝置
3. **多指手勢**：預設僅支援單指操作，多指觸控時不會觸發捲動

## 疑難排解

### 無法偵測到觸控板

**原因**：觸控板未使用 Precision Touchpad 驅動

**解決方法**：
1. 開啟「裝置管理員」
2. 展開「人機介面裝置」
3. 尋找觸控板裝置（通常包含 "HID-compliant touch pad" 或廠商名稱）
4. 右鍵點選 → 內容 → 驅動程式
5. 如果顯示廠商驅動（Synaptics/ELAN），嘗試更新為 Microsoft 標準驅動

### 滑鼠鉤子安裝失敗

**原因**：未以管理員權限執行

**解決方法**：
1. 右鍵點選 `TouchpadSideScroll.exe`
2. 選擇「以系統管理員身分執行」

### 捲動不靈敏或太敏感

**解決方法**：
1. 調整「捲動速度」滑桿
2. 調整「捲動靈敏度」滑桿
3. 嘗試不同的捲動區寬度

## 開發者資訊

### 建置需求

- Visual Studio 2022 或更新版本
- .NET 8.0 SDK
- Windows 10 SDK

### 相依套件

- ModernWpfUI 0.9.6
- Hardcodet.NotifyIcon.Wpf 1.1.0
- Microsoft.Extensions.DependencyInjection 8.0.0
- Serilog 3.1.1

### 參考資源

- [Windows Precision Touchpad Implementation Guide](https://learn.microsoft.com/en-us/windows-hardware/design/component-guidelines/touchpad-windows-precision-touchpad-implementation-guide)
- [Raw Input API](https://learn.microsoft.com/en-us/windows/win32/inputdev/raw-input)
- [HID Application Programming Interface](https://learn.microsoft.com/en-us/windows-hardware/drivers/hid/)

## 授權條款

本專案使用 MIT 授權條款。詳見 [LICENSE](LICENSE) 檔案。

## 致謝

- 使用 [ModernWpf](https://github.com/Kinnara/ModernWpf) 提供現代化 UI
- 使用 [Hardcodet.NotifyIcon.Wpf](https://github.com/hardcodet/wpf-notifyicon) 提供系統匣功能
- 使用 Claude Code 協助開發

## 貢獻

歡迎提交 Issue 和 Pull Request！

---

**注意**：本專案使用 Precision Touchpad API，需要觸控板支援 Windows Precision Touchpad 標準。如果您的觸控板使用原廠驅動程式，可能需要切換到 Microsoft 標準驅動才能使用。
