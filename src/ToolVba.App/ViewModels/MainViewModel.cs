using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using ToolVba.Core;

namespace ToolVba.App.ViewModels;

internal sealed class MainViewModel : INotifyPropertyChanged
{
    private readonly Window _owner;
    private readonly FileOperationService _binaryOperations = new();
    private readonly ExcelModuleInsertionService _moduleOperations = new(auditLog: new SafeOperationAuditLog());
    private readonly ExcelVbaInjector _inspector = new();
    private readonly BasModuleManifestValidator _moduleValidator = new();
    private string _workbookPath = "";
    private string _vbaPath = "";
    private string _modulePath = "";
    private string _manifestPath = "";
    private string _outputPath = "";
    private string _status = "Chọn chế độ và các file đầu vào.";
    private bool _hasError;
    private bool _isWorking;
    private bool _confirmVbaReplacement;
    private int _selectedModeIndex;
    private int _progressPercent;
    private string? _successfulOutput;
    private IReadOnlyList<string> _warnings = [];

    public MainViewModel(Window owner)
    {
        _owner = owner;
        BrowseWorkbookCommand = new RelayCommand(BrowseWorkbook, () => IsIdle);
        BrowseBinaryCommand = new RelayCommand(BrowseBinary, () => IsIdle);
        BrowseModuleCommand = new RelayCommand(BrowseModule, () => IsIdle);
        BrowseManifestCommand = new RelayCommand(BrowseManifest, () => IsIdle);
        BrowseOutputCommand = new RelayCommand(BrowseOutput, () => IsIdle);
        CreateCommand = new RelayCommand(async () => await CreateAsync(), CanCreate);
        OpenOutputFolderCommand = new RelayCommand(OpenOutputFolder, () => CanOpenOutputFolder);
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    public ICommand BrowseWorkbookCommand { get; }
    public ICommand BrowseBinaryCommand { get; }
    public ICommand BrowseModuleCommand { get; }
    public ICommand BrowseManifestCommand { get; }
    public ICommand BrowseOutputCommand { get; }
    public ICommand CreateCommand { get; }
    public ICommand OpenOutputFolderCommand { get; }

    public string WorkbookPath { get => _workbookPath; set { if (Set(ref _workbookPath, value)) RefreshDisplay(nameof(WorkbookDisplay)); } }
    public string VbaPath { get => _vbaPath; set { if (Set(ref _vbaPath, value)) RefreshDisplay(nameof(VbaDisplay)); } }
    public string ModulePath { get => _modulePath; set { if (Set(ref _modulePath, value)) RefreshDisplay(nameof(ModuleDisplay)); } }
    public string ManifestPath { get => _manifestPath; set { if (Set(ref _manifestPath, value)) RefreshDisplay(nameof(ManifestDisplay)); } }
    public string OutputPath { get => _outputPath; set { if (Set(ref _outputPath, value)) RefreshDisplay(nameof(OutputDisplay)); } }
    public string Status { get => _status; private set => Set(ref _status, value); }
    public bool HasError { get => _hasError; private set => Set(ref _hasError, value); }
    public bool IsWorking { get => _isWorking; private set { if (Set(ref _isWorking, value)) { OnPropertyChanged(nameof(IsIdle)); RaiseCommands(); } } }
    public bool IsIdle => !IsWorking;
    public bool ConfirmVbaReplacement { get => _confirmVbaReplacement; set => Set(ref _confirmVbaReplacement, value); }
    public int SelectedModeIndex { get => _selectedModeIndex; set { if (Set(ref _selectedModeIndex, value)) { OnPropertyChanged(nameof(IsModuleMode)); RaiseCommands(); } } }
    public bool IsModuleMode => SelectedModeIndex == 0;
    public int ProgressPercent { get => _progressPercent; private set => Set(ref _progressPercent, value); }
    public string WorkbookDisplay => FileDisplay(WorkbookPath, "Chưa chọn workbook Excel");
    public string VbaDisplay => FileDisplay(VbaPath, "Chưa chọn vbaProject.bin");
    public string ModuleDisplay => FileDisplay(ModulePath, "Chưa chọn module .bas");
    public string ManifestDisplay => FileDisplay(ManifestPath, "Chưa chọn manifest .json");
    public string OutputDisplay => string.IsNullOrWhiteSpace(OutputPath) ? "Chưa chọn nơi lưu" : File.Exists(OutputPath) ? $"{FileDisplay(OutputPath, OutputPath)}  •  {OutputPath}" : OutputPath;
    public bool CanOpenOutputFolder => !string.IsNullOrWhiteSpace(_successfulOutput) && File.Exists(_successfulOutput);
    public IReadOnlyList<string> Warnings { get => _warnings; private set => Set(ref _warnings, value); }

    private void BrowseWorkbook()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Chọn workbook Excel nguồn",
            Filter = "Workbook Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(_owner) != true) return;
        WorkbookPath = dialog.FileName;
        if (string.IsNullOrWhiteSpace(OutputPath))
            OutputPath = Path.Combine(Path.GetDirectoryName(dialog.FileName)!, Path.GetFileNameWithoutExtension(dialog.FileName) + (IsModuleMode ? "_with_module.xlsm" : "_macro.xlsm"));
        ResetSuccess("Đã chọn workbook. File nguồn sẽ không bị sửa.");
    }

    private void BrowseBinary()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Chọn VBA project đã được phê duyệt",
            Filter = "VBA project (*.bin)|*.bin",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(_owner) != true) return;
        VbaPath = dialog.FileName;
        ResetSuccess("Đã chọn VBA project binary.");
    }

    private void BrowseModule()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Chọn module VBA đã được phê duyệt",
            Filter = "Module VBA (*.bas)|*.bas",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(_owner) != true) return;
        ModulePath = dialog.FileName;
        ResetSuccess("Đã chọn module .bas. Công cụ sẽ xác minh hash theo manifest.");
    }

    private void BrowseManifest()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Chọn manifest xác nhận module",
            Filter = "Manifest JSON (*.json)|*.json",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(_owner) != true) return;
        ManifestPath = dialog.FileName;
        ResetSuccess("Đã chọn manifest module.");
    }

    private void BrowseOutput()
    {
        var dialog = new SaveFileDialog
        {
            Title = "Chọn nơi lưu workbook kết quả",
            Filter = "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm",
            DefaultExt = ".xlsm",
            AddExtension = true,
            OverwritePrompt = false,
            FileName = string.IsNullOrWhiteSpace(WorkbookPath) ? "Workbook_output.xlsm" : Path.GetFileNameWithoutExtension(WorkbookPath) + (IsModuleMode ? "_with_module.xlsm" : "_macro.xlsm")
        };
        if (dialog.ShowDialog(_owner) != true) return;
        OutputPath = dialog.FileName;
        ResetSuccess("Đã chọn nơi lưu file kết quả.");
    }

    private bool CanCreate() => IsIdle &&
        File.Exists(WorkbookPath) &&
        !string.IsNullOrWhiteSpace(OutputPath) &&
        (IsModuleMode
            ? File.Exists(ModulePath) && File.Exists(ManifestPath)
            : File.Exists(VbaPath));

    private async Task CreateAsync()
    {
        if (IsModuleMode)
            await CreateWithModuleAsync();
        else
            await CreateWithBinaryAsync();
    }

    private async Task CreateWithModuleAsync()
    {
        var preflight = await Task.Run(() =>
        {
            var sourceInspection = _inspector.InspectWorkbook(WorkbookPath);
            var moduleValidation = _moduleValidator.Validate(ModulePath, ManifestPath);
            return (Source: sourceInspection, Module: moduleValidation);
        });
        var source = preflight.Source;
        if (!source.IsValid)
        {
            SetFailure(source.Error ?? "Chế độ này cần workbook .xlsx hoặc .xlsm hợp lệ.");
            return;
        }

        var moduleCheck = preflight.Module;
        if (!moduleCheck.IsValid)
        {
            SetFailure(moduleCheck.Error ?? "Module không vượt qua kiểm tra manifest.");
            return;
        }

        var summary = $"Nguồn: {Path.GetFileName(WorkbookPath)}\nModule: {Path.GetFileName(ModulePath)}\nKết quả: {OutputPath}\n\nExcel sẽ mở bản tạm với macro bị vô hiệu, import module và lưu output .xlsm riêng. Các component cũ được kiểm tra sau khi lưu. File nguồn không bị sửa. Tool không bật Trust Center hoặc Registry. Tiếp tục?";
        if (MessageBox.Show(_owner, summary, "Xác nhận thêm module", MessageBoxButton.YesNo, MessageBoxImage.Information) != MessageBoxResult.Yes)
            return;

        if (source.HasSignatureArtifacts && MessageBox.Show(_owner,
                "Nguồn có artifact chữ ký package hoặc VBA. Thêm module có thể làm chữ ký cũ mất hiệu lực; công cụ không xác minh chữ ký. Tiếp tục?",
                "Cảnh báo chữ ký", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;

        var overwrite = AskOutputOverwrite();
        if (overwrite is null) return;
        await RunOperationAsync(async progress =>
        {
            var request = new ExcelModuleInsertionRequest(WorkbookPath, ModulePath, ManifestPath, OutputPath, overwrite.Value);
            var result = await _moduleOperations.InsertAsync(request, progress);
            return (result.Succeeded, result.Message, result.OutputPath, result.Warnings);
        }, "Thành công: module đã được thêm vào bản sao; workbook nguồn được giữ nguyên.");
    }

    private async Task CreateWithBinaryAsync()
    {
        var summary = $"Nguồn: {Path.GetFileName(WorkbookPath)}\nVBA: {Path.GetFileName(VbaPath)}\nKết quả: {OutputPath}\n\nFile nguồn sẽ không bị thay đổi. Tiếp tục?";
        if (MessageBox.Show(_owner, summary, "Xác nhận tạo file", MessageBoxButton.YesNo, MessageBoxImage.Information) != MessageBoxResult.Yes)
            return;

        var inspection = await Task.Run(() => _inspector.InspectWorkbook(WorkbookPath));
        if (!inspection.IsValid)
        {
            SetFailure(inspection.Error ?? "File Excel nguồn không hợp lệ.");
            return;
        }
        if (inspection.HasVbaProject && !ConfirmVbaReplacement)
        {
            SetFailure("Workbook đã có VBA project. Chọn tùy chọn thay thế để tiếp tục.");
            return;
        }
        if (inspection.HasVbaProject && MessageBox.Show(_owner,
                "Thao tác này thay toàn bộ VBA project trong bản kết quả; các macro cũ sẽ không được ghép tự động. Chỉ tiếp tục nếu binary mới đã chứa nội dung cần giữ.",
                "Xác nhận thay toàn bộ VBA", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;
        if (inspection.HasSignatureArtifacts && MessageBox.Show(_owner,
                "Nguồn có artifact chữ ký package hoặc VBA. Việc sửa package có thể làm chữ ký cũ mất hiệu lực. Tiếp tục?",
                "Cảnh báo chữ ký", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;

        var overwrite = AskOutputOverwrite();
        if (overwrite is null) return;
        await RunOperationAsync(async progress =>
        {
            var request = new InjectionRequest(WorkbookPath, VbaPath, OutputPath, ConfirmVbaReplacement, overwrite.Value);
            var result = await _binaryOperations.CreateAsync(request, new Progress<OperationProgress>(p => progress.Report(new(p.Stage, p.Percent))));
            return (result.Succeeded, result.Message, result.OutputPath, result.Warnings);
        }, "Thành công: file kết quả đã vượt qua kiểm tra package và SHA-256.");
    }

    private bool? AskOutputOverwrite()
    {
        if (!File.Exists(OutputPath)) return false;
        return MessageBox.Show(_owner, $"File kết quả đã tồn tại:\n{OutputPath}\n\nChỉ file này sẽ được thay sau khi toàn bộ kiểm tra đạt. Ghi đè?",
            "Xác nhận ghi đè", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes ? true : null;
    }

    private async Task RunOperationAsync(
        Func<IProgress<ExcelModuleOperationProgress>, Task<(bool Succeeded, string Message, string? OutputPath, IReadOnlyList<string>? Warnings)>> operation,
        string successMessage)
    {
        Warnings = [];
        HasError = false;
        _successfulOutput = null;
        OnPropertyChanged(nameof(CanOpenOutputFolder));
        ProgressPercent = 0;
        IsWorking = true;
        Status = "Đang xử lý trên bản tạm…";
        try
        {
            var progress = new Progress<ExcelModuleOperationProgress>(p =>
            {
                Status = p.Stage;
                ProgressPercent = p.Percent;
            });
            var result = await operation(progress);
            if (!result.Succeeded)
            {
                SetFailure(result.Message);
                return;
            }
            _successfulOutput = result.OutputPath;
            Warnings = result.Warnings ?? [];
            Status = successMessage;
            OnPropertyChanged(nameof(CanOpenOutputFolder));
        }
        catch (Exception)
        {
            SetFailure("Đã xảy ra lỗi khi xử lý. Kiểm tra quyền file/Excel và thử lại trên bản sao.");
        }
        finally
        {
            IsWorking = false;
        }
    }

    private void OpenOutputFolder()
    {
        if (!CanOpenOutputFolder) return;
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{_successfulOutput}\"") { UseShellExecute = true });
        }
        catch (System.ComponentModel.Win32Exception)
        {
            SetFailure("Không thể mở Explorer. File kết quả vẫn ở đường dẫn đã chọn.");
        }
    }

    private void ResetSuccess(string message)
    {
        HasError = false;
        Status = message;
        _successfulOutput = null;
        OnPropertyChanged(nameof(CanOpenOutputFolder));
        RaiseCommands();
    }

    private void SetFailure(string message)
    {
        HasError = true;
        Status = "Thất bại: " + message;
        Warnings = [];
        _successfulOutput = null;
        OnPropertyChanged(nameof(CanOpenOutputFolder));
    }

    private static string FileDisplay(string path, string fallback)
    {
        if (string.IsNullOrWhiteSpace(path)) return fallback;
        try
        {
            var file = new FileInfo(path);
            return file.Exists ? $"{file.Name}  •  {FormatSize(file.Length)}" : file.Name;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return Path.GetFileName(path);
        }
    }

    private static string FormatSize(long bytes) => bytes switch
    {
        < 1024 => $"{bytes} B",
        < 1024 * 1024 => $"{bytes / 1024d:0.0} KB",
        < 1024L * 1024 * 1024 => $"{bytes / (1024d * 1024):0.0} MB",
        _ => $"{bytes / (1024d * 1024 * 1024):0.00} GB"
    };

    private void RefreshDisplay(string propertyName)
    {
        OnPropertyChanged(propertyName);
        RaiseCommands();
    }

    private void RaiseCommands()
    {
        (BrowseWorkbookCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (BrowseBinaryCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (BrowseModuleCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (BrowseManifestCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (BrowseOutputCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (CreateCommand as RelayCommand)?.RaiseCanExecuteChanged();
        (OpenOutputFolderCommand as RelayCommand)?.RaiseCanExecuteChanged();
    }

    private bool Set<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
