using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using ToolVba.App.Services;
using ToolVba.Core;

namespace ToolVba.App.ViewModels;

/// <summary>Single screen: drop Excel files, press Create, the company VBA package is pasted into copies.</summary>
internal sealed class MainViewModel : ObservableObject
{
    private readonly Window _owner;
    private readonly AppServices _services;
    private string? _outputDirectory;
    private string _status = "Kéo thả file Excel vào cửa sổ này, rồi bấm Tạo.";
    private bool _hasError;
    private bool _isWorking;
    private int _progressPercent;
    private string? _lastOutput;
    private IReadOnlyList<string> _warnings = [];

    public MainViewModel(Window owner, AppServices services)
    {
        _owner = owner;
        _services = services;
        Package = new VbaPackageStatusViewModel(services.Packages, services.ConfigurationError);
        Update = new AppUpdateViewModel(services.Updater, () => IsWorking);
        Trust = new TrustBannerViewModel(services.Trust, owner);
        BrowseCommand = new RelayCommand(Browse, () => IsIdle);
        CreateCommand = new RelayCommand(async () => await CreateAsync(), CanCreate);
        ClearCommand = new RelayCommand(Files.Clear, () => IsIdle && HasFiles);
        ChangeOutputFolderCommand = new RelayCommand(ChangeOutputFolder, () => IsIdle);
        ResetOutputFolderCommand = new RelayCommand(() => SetOutputDirectory(null), () => IsIdle && _outputDirectory is not null);
        OpenOutputCommand = new RelayCommand(OpenOutput, () => CanOpenOutput);
        OpenFolderCommand = new RelayCommand(OpenFolder, () => CanOpenOutput);
        Files.CollectionChanged += (_, _) => { OnPropertyChanged(nameof(HasFiles)); OnPropertyChanged(nameof(CreateButtonText)); RaiseCommands(); };
        Package.PropertyChanged += (_, _) => RaiseCommands();
        Update.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(AppUpdateViewModel.IsDownloading)) RaiseCommands(); };
    }

    /// <summary>True while work is running that closing the window or restarting would interrupt.</summary>
    public bool IsBusy => IsWorking || Update.IsDownloading;

    public VbaPackageStatusViewModel Package { get; }
    public AppUpdateViewModel Update { get; }
    public TrustBannerViewModel Trust { get; }
    public ObservableCollection<WorkbookItem> Files { get; } = [];

    public ICommand BrowseCommand { get; }
    public ICommand CreateCommand { get; }
    public ICommand ClearCommand { get; }
    public ICommand ChangeOutputFolderCommand { get; }
    public ICommand ResetOutputFolderCommand { get; }
    public ICommand OpenOutputCommand { get; }
    public ICommand OpenFolderCommand { get; }

    public string WindowTitle => $"Đóng gói VBA — phiên bản {AppConfig.AppVersionText}";
    public bool HasFiles => Files.Count > 0;
    public string OutputFolderText => _outputDirectory ?? "Lưu cạnh file gốc";
    public bool HasCustomOutputFolder => _outputDirectory is not null;
    public string CreateButtonText => Files.Count > 1 ? $"Tạo {Files.Count} file có VBA" : "Tạo file có VBA";
    public string Status { get => _status; private set => Set(ref _status, value); }
    public bool HasError { get => _hasError; private set => Set(ref _hasError, value); }
    public int ProgressPercent { get => _progressPercent; private set => Set(ref _progressPercent, value); }
    public IReadOnlyList<string> Warnings { get => _warnings; private set => Set(ref _warnings, value); }
    public bool IsIdle => !IsWorking;
    public bool CanOpenOutput => IsIdle && _lastOutput is not null && File.Exists(_lastOutput);

    public bool IsWorking
    {
        get => _isWorking;
        private set
        {
            if (!Set(ref _isWorking, value)) return;
            OnPropertyChanged(nameof(IsIdle));
            OnPropertyChanged(nameof(CanOpenOutput));
            RaiseCommands();
        }
    }

    public async Task InitializeAsync()
    {
        Trust.Refresh();
        var status = await Package.RefreshAsync();
        Update.RequireVersion(status.RequiredAppVersion);
        await Update.StartAsync();
    }

    public static bool IsSupportedFile(string path) =>
        Path.GetExtension(path) is { } extension &&
        (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) || extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase));

    public void AddFiles(IEnumerable<string> paths)
    {
        if (IsWorking) return;
        if (Files.Count > 0 && Files.All(f => f.IsFinished)) Files.Clear();
        var skipped = 0;
        foreach (var path in paths)
        {
            if (!IsSupportedFile(path) || !File.Exists(path)) { skipped++; continue; }
            if (Files.Any(f => string.Equals(f.SourcePath, path, StringComparison.OrdinalIgnoreCase))) continue;
            Files.Add(new WorkbookItem(path, ResolveOutput(path, null), item => Files.Remove(item)));
        }
        HasError = false;
        Status = skipped > 0
            ? $"Đã bỏ qua {skipped} file không phải Excel .xlsx/.xlsm."
            : "Bấm Tạo để dán code VBA. File gốc được giữ nguyên.";
    }

    // Blocked while an app update downloads: applying it restarts the app and would cut an install short.
    private bool CanCreate() => IsIdle && !Update.IsDownloading && Files.Any(f => f.State != WorkbookItemState.Done) && Package.Current is not null;

    private async Task CreateAsync()
    {
        IsWorking = true;
        HasError = false;
        Warnings = [];
        _lastOutput = null;
        try
        {
            var package = await PrepareAsync();
            if (package is null) return;
            await InstallAllAsync(package);
        }
        catch (Exception ex)
        {
            // UI boundary: the install service already turns expected failures into results.
            AppErrorLog.Write("create", ex);
            SetFailure("Đã xảy ra lỗi khi xử lý. Hãy đóng hết Excel rồi thử lại.");
        }
        finally
        {
            IsWorking = false;
        }
    }

    /// <summary>Makes sure the newest verified package is used and Excel allows VBA project access.</summary>
    private async Task<VbaPackage?> PrepareAsync()
    {
        Status = "Đang kiểm tra code VBA mới nhất…";
        var status = await Package.RefreshAsync();
        Update.RequireVersion(status.RequiredAppVersion);
        if (!status.CanInstall)
        {
            SetFailure(status.Message);
            return null;
        }
        if (Trust.Refresh() != VbaTrustState.Enabled)
        {
            SetFailure(Trust.Message + (Trust.CanEnable ? " Bấm \"Bật giúp tôi\" ở phía trên rồi bấm Tạo lại." : ""));
            return null;
        }
        return status.Package;
    }

    private async Task InstallAllAsync(VbaPackage package)
    {
        var pending = Files.Where(f => f.State != WorkbookItemState.Done).ToList();
        var warnings = new List<string>();
        var succeeded = 0;
        for (var index = 0; index < pending.Count; index++)
        {
            var item = pending[index];
            var position = index;
            item.State = WorkbookItemState.Working;
            item.Message = null;
            item.OutputPath = ResolveOutput(item.SourcePath, item);
            var progress = new Progress<VbaInstallProgress>(p =>
            {
                Status = pending.Count > 1 ? $"[{position + 1}/{pending.Count}] {item.FileName}: {p.Stage}" : p.Stage;
                ProgressPercent = (position * 100 + p.Percent) / pending.Count;
            });

            VbaPackageInstallResult result;
            try
            {
                result = await _services.Installer.InstallAsync(new(item.SourcePath, item.OutputPath, package), progress);
            }
            finally
            {
                // Never leave a row stuck on "Đang xử lý…" (its remove button stays disabled in that state).
                if (item.State == WorkbookItemState.Working) item.State = WorkbookItemState.Failed;
            }
            item.State = result.Succeeded ? WorkbookItemState.Done : WorkbookItemState.Failed;
            item.Message = result.Succeeded ? null : result.Message;
            if (!result.Succeeded) continue;
            succeeded++;
            _lastOutput = result.OutputPath;
            warnings.AddRange(result.Warnings ?? []);
        }

        Warnings = warnings.Distinct().ToList();
        OnPropertyChanged(nameof(CanOpenOutput));
        ReportBatch(pending, succeeded, package.Version);
    }

    private void ReportBatch(IReadOnlyList<WorkbookItem> processed, int succeeded, string version)
    {
        ProgressPercent = 100;
        if (succeeded == processed.Count)
        {
            Status = processed.Count == 1
                ? $"Xong! Đã tạo {processed[0].OutputName} với code VBA v{version}. File gốc được giữ nguyên."
                : $"Xong! Đã tạo {succeeded} file với code VBA v{version}. File gốc được giữ nguyên.";
            return;
        }
        Trust.Refresh();
        SetFailure(processed.Count == 1
            ? processed[0].Message ?? "Không tạo được file."
            : $"Đã tạo {succeeded}/{processed.Count} file. Xem lý do ở từng dòng báo lỗi, sửa rồi bấm Tạo lại.");
    }

    private string ResolveOutput(string source, WorkbookItem? self)
    {
        var reserved = Files.Where(f => !ReferenceEquals(f, self)).Select(f => f.OutputPath)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return OutputPathResolver.Resolve(source, _outputDirectory, reserved);
    }

    private void Browse()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Chọn file Excel cần dán code VBA",
            Filter = "File Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            CheckFileExists = true,
            Multiselect = true
        };
        if (dialog.ShowDialog(_owner) == true) AddFiles(dialog.FileNames);
    }

    private void ChangeOutputFolder()
    {
        var dialog = new OpenFolderDialog { Title = "Chọn thư mục lưu file kết quả" };
        if (dialog.ShowDialog(_owner) == true) SetOutputDirectory(dialog.FolderName);
    }

    private void SetOutputDirectory(string? directory)
    {
        _outputDirectory = directory;
        foreach (var item in Files.Where(f => f.State != WorkbookItemState.Done))
            item.OutputPath = ResolveOutput(item.SourcePath, item);
        OnPropertyChanged(nameof(OutputFolderText));
        OnPropertyChanged(nameof(HasCustomOutputFolder));
        RaiseCommands();
    }

    private void OpenOutput() => Launch(new ProcessStartInfo(_lastOutput!) { UseShellExecute = true });

    private void OpenFolder() => Launch(new ProcessStartInfo("explorer.exe", $"/select,\"{_lastOutput}\"") { UseShellExecute = true });

    private void Launch(ProcessStartInfo startInfo)
    {
        if (!CanOpenOutput) return;
        try
        {
            Process.Start(startInfo);
        }
        catch (System.ComponentModel.Win32Exception)
        {
            SetFailure("Không mở được file hoặc thư mục. File kết quả vẫn nằm ở: " + _lastOutput);
        }
    }

    private void SetFailure(string message)
    {
        HasError = true;
        Status = message;
    }

    private void RaiseCommands()
    {
        foreach (var command in new[] { BrowseCommand, CreateCommand, ClearCommand, ChangeOutputFolderCommand,
                     ResetOutputFolderCommand, OpenOutputCommand, OpenFolderCommand })
            (command as RelayCommand)?.RaiseCanExecuteChanged();
        Update.RaiseCanExecute();
    }
}
