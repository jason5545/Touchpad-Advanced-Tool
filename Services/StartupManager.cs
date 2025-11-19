using System;
using System.IO;
using System.Reflection;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;

namespace TouchpadAdvancedTool.Services
{
    /// <summary>
    /// 啟動管理器 - 負責管理應用程式的開機自動啟動功能
    /// </summary>
    public class StartupManager
    {
        private readonly ILogger<StartupManager> _logger;
        private const string RegistryKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string AppName = "TouchpadAdvancedTool";

        public StartupManager(ILogger<StartupManager> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// 啟用開機自動啟動
        /// </summary>
        /// <param name="silentStart">是否靜默啟動（啟動到系統匣）</param>
        /// <returns>是否成功</returns>
        public bool EnableStartup(bool silentStart = true)
        {
            try
            {
                var exePath = GetExecutablePath();
                if (string.IsNullOrEmpty(exePath))
                {
                    _logger.LogError("無法取得執行檔路徑");
                    return false;
                }

                // 建構啟動命令
                var command = silentStart ? $"\"{exePath}\" --minimized" : $"\"{exePath}\"";

                using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, true);
                if (key == null)
                {
                    _logger.LogError("無法開啟註冊表鍵：{Path}", RegistryKeyPath);
                    return false;
                }

                key.SetValue(AppName, command, RegistryValueKind.String);
                _logger.LogInformation("已啟用開機自動啟動：{Command}", command);
                return true;
            }
            catch (UnauthorizedAccessException ex)
            {
                _logger.LogError(ex, "啟用開機自動啟動失敗：權限不足");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "啟用開機自動啟動失敗");
                return false;
            }
        }

        /// <summary>
        /// 停用開機自動啟動
        /// </summary>
        /// <returns>是否成功</returns>
        public bool DisableStartup()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, true);
                if (key == null)
                {
                    _logger.LogError("無法開啟註冊表鍵：{Path}", RegistryKeyPath);
                    return false;
                }

                if (key.GetValue(AppName) != null)
                {
                    key.DeleteValue(AppName);
                    _logger.LogInformation("已停用開機自動啟動");
                }
                return true;
            }
            catch (UnauthorizedAccessException ex)
            {
                _logger.LogError(ex, "停用開機自動啟動失敗：權限不足");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "停用開機自動啟動失敗");
                return false;
            }
        }

        /// <summary>
        /// 檢查是否已啟用開機自動啟動
        /// </summary>
        /// <returns>是否已啟用</returns>
        public bool IsStartupEnabled()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, false);
                var value = key?.GetValue(AppName) as string;
                return !string.IsNullOrEmpty(value);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "檢查開機自動啟動狀態失敗");
                return false;
            }
        }

        /// <summary>
        /// 取得當前註冊的啟動命令
        /// </summary>
        /// <returns>啟動命令，如果未啟用則返回 null</returns>
        public string? GetStartupCommand()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, false);
                return key?.GetValue(AppName) as string;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "取得啟動命令失敗");
                return null;
            }
        }

        /// <summary>
        /// 驗證啟動項目是否指向正確的執行檔
        /// </summary>
        /// <returns>是否有效</returns>
        public bool ValidateStartupEntry()
        {
            try
            {
                var command = GetStartupCommand();
                if (string.IsNullOrEmpty(command))
                    return false;

                var currentExePath = GetExecutablePath();
                if (string.IsNullOrEmpty(currentExePath))
                    return false;

                // 檢查命令中是否包含當前執行檔路徑
                return command.Contains(currentExePath, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "驗證啟動項目失敗");
                return false;
            }
        }

        /// <summary>
        /// 取得執行檔路徑
        /// </summary>
        /// <returns>執行檔路徑</returns>
        private string? GetExecutablePath()
        {
            try
            {
                // 取得當前執行檔路徑
                return System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "取得執行檔路徑失敗");
                return null;
            }
        }
    }
}

