using System.IO;
using ToolVba.Core;

namespace ToolVba.App.ViewModels;

/// <summary>The small chip that shows which company VBA version will be pasted.</summary>
internal sealed class VbaPackageStatusViewModel(VbaPackageService? service, string? configurationError) : ObservableObject
{
    private Task<VbaPackageStatus>? _inFlight;
    private VbaPackageStatus? _current;
    private string _text = "Đang kiểm tra code VBA…";
    private bool _isChecking;

    public VbaPackageStatus? Current { get => _current; private set { if (Set(ref _current, value)) RaiseDerived(); } }
    public string Text { get => _text; private set => Set(ref _text, value); }
    public bool IsChecking { get => _isChecking; private set => Set(ref _isChecking, value); }
    public bool IsReady => Current?.CanInstall == true;
    public bool IsLatest => IsReady && Current!.Availability == VbaPackageAvailability.Latest;
    public bool HasProblem => !IsReady;

    /// <summary>Asks the server for the latest package; concurrent callers share one request.</summary>
    public Task<VbaPackageStatus> RefreshAsync()
    {
        if (_inFlight is { IsCompleted: false }) return _inFlight;
        _inFlight = RefreshCoreAsync();
        return _inFlight;
    }

    private async Task<VbaPackageStatus> RefreshCoreAsync()
    {
        IsChecking = true;
        try
        {
            if (service is null)
            {
                Current = VbaPackageStatus.Unavailable(configurationError ?? "Chưa cấu hình máy chủ cập nhật.");
            }
            else
            {
                try
                {
                    Current = await service.RefreshAsync();
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    Current = VbaPackageStatus.Unavailable("Không ghi được thư mục lưu code VBA trên máy (%LocalAppData%\\ToolVba).");
                }
            }
            Text = Current.Message;
            return Current;
        }
        finally
        {
            IsChecking = false;
        }
    }

    private void RaiseDerived()
    {
        OnPropertyChanged(nameof(IsReady));
        OnPropertyChanged(nameof(IsLatest));
        OnPropertyChanged(nameof(HasProblem));
    }
}
