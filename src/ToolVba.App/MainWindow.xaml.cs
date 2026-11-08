using System.Windows;
using ToolVba.App.Services;
using ToolVba.App.ViewModels;

namespace ToolVba.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    internal MainWindow(AppServices services)
    {
        InitializeComponent();
        _viewModel = new MainViewModel(this, services);
        DataContext = _viewModel;
        Loaded += async (_, _) =>
        {
            try
            {
                // Files dropped onto the shortcut or opened with "Open with" arrive as arguments.
                var files = Environment.GetCommandLineArgs().Skip(1)
                    .Where(path => MainViewModel.IsSupportedFile(path) && System.IO.File.Exists(path)).ToList();
                if (files.Count > 0) _viewModel.AddFiles(files);
                await _viewModel.InitializeAsync();
            }
            catch (Exception ex)
            {
                // async void handler: keep the window usable even if the first check fails.
                AppErrorLog.Write("startup", ex);
            }
        };
        Closing += OnClosing;
    }

    /// <summary>
    /// Closing mid-install would kill the Excel worker thread and leave a hidden EXCEL.EXE and a temporary
    /// copy next to the output, so the window stays open until the work finishes.
    /// </summary>
    private void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!_viewModel.IsBusy) return;
        e.Cancel = true;
        MessageBox.Show(this, "Ứng dụng đang dán code hoặc đang tải bản cập nhật. Hãy đợi xong rồi đóng.",
            "Đóng gói VBA", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OnDragEnter(object sender, DragEventArgs e)
    {
        UpdateDragEffect(e);
        DropOverlay.Visibility = e.Effects == DragDropEffects.Copy ? Visibility.Visible : Visibility.Collapsed;
    }

    private void OnDragOver(object sender, DragEventArgs e) => UpdateDragEffect(e);

    private void OnDragLeave(object sender, DragEventArgs e) => DropOverlay.Visibility = Visibility.Collapsed;

    private void OnDrop(object sender, DragEventArgs e)
    {
        DropOverlay.Visibility = Visibility.Collapsed;
        if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
            _viewModel.AddFiles(files);
        e.Handled = true;
    }

    private void UpdateDragEffect(DragEventArgs e)
    {
        var hasExcel = _viewModel.IsIdle &&
                       e.Data.GetData(DataFormats.FileDrop) is string[] files &&
                       files.Any(MainViewModel.IsSupportedFile);
        e.Effects = hasExcel ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }
}
