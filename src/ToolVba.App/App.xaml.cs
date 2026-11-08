using System.Windows;
using System.Windows.Threading;
using ToolVba.App.Services;

namespace ToolVba.App;

public partial class App : Application
{
    private AppServices? _services;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            AppErrorLog.Write("unobserved-task", args.Exception);
            args.SetObserved();
        };
        _services = AppServices.Create();
        var window = new MainWindow(_services);
        MainWindow = window;
        window.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _services?.Dispose();
        base.OnExit(e);
    }

    /// <summary>Last line of defence: keep the app open and tell the user instead of crashing.</summary>
    private void OnUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        AppErrorLog.Write("dispatcher", e.Exception);
        e.Handled = true;
        MessageBox.Show(MainWindow, "Đã xảy ra lỗi không mong muốn. Ứng dụng vẫn chạy; nếu lỗi lặp lại, hãy đóng hết Excel và mở lại ứng dụng.",
            "Đóng gói VBA", MessageBoxButton.OK, MessageBoxImage.Warning);
    }
}
