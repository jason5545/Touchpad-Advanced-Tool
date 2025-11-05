using System;
using System.Threading;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;
using TouchpadSideScroll.Core;
using TouchpadSideScroll.Services;
using TouchpadSideScroll.ViewModels;

namespace TouchpadSideScroll
{
    /// <summary>
    /// 應用程式主類別
    /// </summary>
    public partial class App : Application
    {
        private static Mutex? _mutex;
        private IServiceProvider? _serviceProvider;

        /// <summary>
        /// 應用程式啟動
        /// </summary>
        protected override void OnStartup(StartupEventArgs e)
        {
            // 檢查是否已有執行個體在執行（單一執行個體）
            const string mutexName = "TouchpadSideScroll_SingleInstance_Mutex";
            _mutex = new Mutex(true, mutexName, out bool createdNew);

            if (!createdNew)
            {
                MessageBox.Show(
                    "觸控板側邊捲動已在執行中。",
                    "TouchpadSideScroll",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                Shutdown();
                return;
            }

            // 設定 Serilog
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(
                    path: System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "TouchpadSideScroll",
                        "logs",
                        "log-.txt"),
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7)
                .CreateLogger();

            // 設定依賴注入
            var services = new ServiceCollection();
            ConfigureServices(services);
            _serviceProvider = services.BuildServiceProvider();

            // 處理未捕獲的例外狀況
            DispatcherUnhandledException += (s, args) =>
            {
                Log.Error(args.Exception, "未處理的例外狀況");
                MessageBox.Show(
                    $"發生錯誤：{args.Exception.Message}",
                    "錯誤",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                args.Handled = true;
            };

            base.OnStartup(e);

            // 顯示主視窗
            var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
            mainWindow.Show();
        }

        /// <summary>
        /// 設定服務
        /// </summary>
        private void ConfigureServices(ServiceCollection services)
        {
            // 日誌服務
            services.AddLogging(builder =>
            {
                builder.AddSerilog(dispose: true);
            });

            // 核心服務
            services.AddSingleton<RawInputManager>();
            services.AddSingleton<MouseHookManager>();
            services.AddSingleton<TouchpadTracker>();
            services.AddSingleton<ScrollConverter>();
            services.AddSingleton<SettingsManager>();

            // ViewModels
            services.AddSingleton<MainViewModel>();

            // Views
            services.AddSingleton<MainWindow>();
        }

        /// <summary>
        /// 應用程式關閉
        /// </summary>
        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                var hook = _serviceProvider?.GetService<MouseHookManager>();
                hook?.UninstallHook();

                var raw = _serviceProvider?.GetService<RawInputManager>();
                raw?.Unregister();

                if (_serviceProvider is IDisposable d)
                {
                    d.Dispose();
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "應用程式關閉時清理失敗");
            }

            Log.Information("應用程式關閉");
            Log.CloseAndFlush();

            _mutex?.ReleaseMutex();
            _mutex?.Dispose();

            base.OnExit(e);
        }

        /// <summary>
        /// 取得服務
        /// </summary>
        public static T GetService<T>() where T : class
        {
            var app = (App)Current;
            return app._serviceProvider?.GetService<T>()
                ?? throw new InvalidOperationException($"服務 {typeof(T).Name} 未註冊");
        }
    }
}
