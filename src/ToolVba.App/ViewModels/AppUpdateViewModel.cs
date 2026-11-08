using System.Windows.Input;
using System.Windows.Threading;
using ToolVba.App.Services;
using Velopack;

namespace ToolVba.App.ViewModels;

/// <summary>Banner that announces a newer app version and installs it with one click.</summary>
internal sealed class AppUpdateViewModel : ObservableObject
{
    private static readonly TimeSpan CheckInterval = TimeSpan.FromHours(4);
    private readonly AppUpdater _updater;
    private readonly Func<bool> _isBusy;
    private readonly DispatcherTimer _timer;
    private UpdateInfo? _update;
    private string? _requiredVersion;
    private bool _isVisible;
    private bool _isDownloading;
    private int _progressPercent;
    private string _title = "";
    private string _detail = "";

    public AppUpdateViewModel(AppUpdater updater, Func<bool> isBusy)
    {
        _updater = updater;
        _isBusy = isBusy;
        UpdateNowCommand = new RelayCommand(async () => await UpdateNowAsync(), CanUpdateNow);
        _timer = new DispatcherTimer { Interval = CheckInterval };
        _timer.Tick += async (_, _) => await CheckAsync();
    }

    public ICommand UpdateNowCommand { get; }
    public bool IsVisible { get => _isVisible; private set => Set(ref _isVisible, value); }
    public bool IsDownloading { get => _isDownloading; private set { if (Set(ref _isDownloading, value)) RaiseCanExecute(); } }
    public int ProgressPercent { get => _progressPercent; private set => Set(ref _progressPercent, value); }
    public string Title { get => _title; private set => Set(ref _title, value); }
    public string Detail { get => _detail; private set => Set(ref _detail, value); }
    public bool CanInstallUpdate => _update is not null;

    public async Task StartAsync()
    {
        _timer.Start();
        await CheckAsync();
    }

    /// <summary>Called when the current VBA package needs a newer app than this one.</summary>
    public void RequireVersion(string? version)
    {
        _requiredVersion = version;
        Refresh();
    }

    public void RaiseCanExecute() => (UpdateNowCommand as RelayCommand)?.RaiseCanExecuteChanged();

    private async Task CheckAsync()
    {
        if (IsDownloading) return;
        var check = await _updater.CheckAsync();
        _update = check.Update;
        Refresh(check.Version, check.Notes);
    }

    private void Refresh(string? version = null, string? notes = null)
    {
        OnPropertyChanged(nameof(CanInstallUpdate));
        RaiseCanExecute();
        if (_update is not null)
        {
            IsVisible = true;
            Title = $"Có phiên bản ứng dụng mới {version ?? _update.TargetFullRelease.Version.ToString()}";
            Detail = _requiredVersion is null
                ? FirstLine(notes) ?? "Bấm Cập nhật ngay, ứng dụng sẽ tự tải và mở lại."
                : "Code VBA mới cần bản này. Bấm Cập nhật ngay, ứng dụng sẽ tự tải và mở lại.";
            return;
        }
        IsVisible = _requiredVersion is not null;
        Title = $"Cần ứng dụng bản {_requiredVersion} trở lên";
        Detail = _updater.IsSupported
            ? "Máy chủ chưa có bản ứng dụng này. Hãy liên hệ quản trị viên."
            : "Bản đang chạy không tự cập nhật được. Hãy cài lại bằng file ToolVba-win-Setup.exe mới nhất.";
    }

    private bool CanUpdateNow() => _update is not null && !IsDownloading && !_isBusy();

    private async Task UpdateNowAsync()
    {
        if (_update is null) return;
        IsDownloading = true;
        var progress = new Progress<int>(p => ProgressPercent = p);
        try
        {
            Detail = "Đang tải bản cập nhật…";
            await _updater.DownloadAsync(_update, p => ((IProgress<int>)progress).Report(p), CancellationToken.None);
            Detail = "Đang cài đặt và mở lại ứng dụng…";
            _updater.ApplyAndRestart(_update);
        }
        catch (Exception ex)
        {
            // UI boundary: Velopack can fail in many ways (network, disk, locked files); tell the user and keep running.
            AppErrorLog.Write("update-download", ex);
            Detail = "Không tải được bản cập nhật. Hãy kiểm tra mạng rồi bấm lại.";
        }
        finally
        {
            IsDownloading = false;
        }
    }

    private static string? FirstLine(string? text) =>
        text?.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(l => l.TrimStart('#', '-', '*', ' '))
            .FirstOrDefault(l => l.Length > 0);
}
