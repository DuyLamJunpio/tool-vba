using System.IO;
using System.Windows.Input;

namespace ToolVba.App.ViewModels;

internal enum WorkbookItemState { Ready, Working, Done, Failed }

/// <summary>One Excel file the user dropped, with where its result goes and how it went.</summary>
internal sealed class WorkbookItem : ObservableObject
{
    private string _outputPath;
    private WorkbookItemState _state;
    private string? _message;

    public WorkbookItem(string path, string outputPath, Action<WorkbookItem> remove)
    {
        SourcePath = path;
        _outputPath = outputPath;
        FileName = Path.GetFileName(path);
        SizeText = SafeSize(path);
        RemoveCommand = new RelayCommand(() => remove(this), () => State != WorkbookItemState.Working);
    }

    public string SourcePath { get; }
    public string FileName { get; }
    public string SizeText { get; }
    public ICommand RemoveCommand { get; }

    public string OutputPath
    {
        get => _outputPath;
        set { if (Set(ref _outputPath, value)) OnPropertyChanged(nameof(OutputName)); }
    }

    public string OutputName => Path.GetFileName(OutputPath);

    public WorkbookItemState State
    {
        get => _state;
        set
        {
            if (!Set(ref _state, value)) return;
            OnPropertyChanged(nameof(StateText));
            OnPropertyChanged(nameof(IsFinished));
            (RemoveCommand as RelayCommand)?.RaiseCanExecuteChanged();
        }
    }

    public string? Message { get => _message; set => Set(ref _message, value); }

    public bool IsFinished => State is WorkbookItemState.Done or WorkbookItemState.Failed;

    public string StateText => State switch
    {
        WorkbookItemState.Working => "Đang xử lý…",
        WorkbookItemState.Done => "Xong",
        WorkbookItemState.Failed => "Lỗi",
        _ => "Sẵn sàng"
    };

    public static string FormatSize(long bytes) => bytes switch
    {
        < 1024 => $"{bytes} B",
        < 1024 * 1024 => $"{bytes / 1024d:0.0} KB",
        < 1024L * 1024 * 1024 => $"{bytes / (1024d * 1024):0.0} MB",
        _ => $"{bytes / (1024d * 1024 * 1024):0.00} GB"
    };

    private static string SafeSize(string path)
    {
        try { return FormatSize(new FileInfo(path).Length); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException) { return ""; }
    }
}
