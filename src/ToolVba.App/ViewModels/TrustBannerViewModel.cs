using System.IO;
using System.Security;
using System.Windows;
using System.Windows.Input;
using ToolVba.Core;

namespace ToolVba.App.ViewModels;

/// <summary>Explains when Excel blocks VBA project access and offers to enable it for this Windows user.</summary>
internal sealed class TrustBannerViewModel : ObservableObject
{
    private readonly ExcelTrustSettings _settings;
    private readonly Window _owner;
    private VbaTrustState _state = VbaTrustState.Enabled;
    private string _message = "";

    public TrustBannerViewModel(ExcelTrustSettings settings, Window owner)
    {
        _settings = settings;
        _owner = owner;
        EnableCommand = new RelayCommand(Enable, () => _state == VbaTrustState.Disabled);
    }

    public ICommand EnableCommand { get; }
    public bool IsVisible => _state != VbaTrustState.Enabled;
    public bool CanEnable => _state == VbaTrustState.Disabled;
    public string Message { get => _message; private set => Set(ref _message, value); }
    public bool IsEnabled => _state == VbaTrustState.Enabled;

    public VbaTrustState Refresh()
    {
        Apply(_settings.GetState().State);
        return _state;
    }

    private void Enable()
    {
        var answer = MessageBox.Show(_owner,
            "Công cụ sẽ bật tùy chọn \"Trust access to the VBA project object model\" của Excel cho riêng tài khoản Windows này, " +
            "để có thể dán code VBA vào file.\n\nTùy chọn này cho phép chương trình trên máy đọc và sửa code VBA. Tiếp tục?",
            "Bật quyền dán code VBA", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (answer != MessageBoxResult.Yes) return;
        try
        {
            Apply(_settings.TryEnableForCurrentUser().State);
            if (IsEnabled)
                MessageBox.Show(_owner, "Đã bật. Nếu Excel đang mở, hãy đóng hết cửa sổ Excel trước khi bấm Tạo.",
                    "Bật quyền dán code VBA", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex) when (ex is UnauthorizedAccessException or SecurityException or IOException)
        {
            Message = "Windows không cho phép ghi thiết lập này. Hãy liên hệ bộ phận IT.";
        }
    }

    private void Apply(VbaTrustState state)
    {
        _state = state;
        Message = state switch
        {
            VbaTrustState.Disabled => "Excel chưa cho phép công cụ dán code VBA vào file.",
            VbaTrustState.BlockedByPolicy => "Chính sách của công ty đang chặn quyền dán code VBA. Hãy liên hệ bộ phận IT.",
            VbaTrustState.ExcelNotFound => "Không tìm thấy Microsoft Excel trên máy này. Cần cài Excel để dán code VBA.",
            _ => ""
        };
        OnPropertyChanged(nameof(IsVisible));
        OnPropertyChanged(nameof(CanEnable));
        OnPropertyChanged(nameof(IsEnabled));
        (EnableCommand as RelayCommand)?.RaiseCanExecuteChanged();
    }
}
