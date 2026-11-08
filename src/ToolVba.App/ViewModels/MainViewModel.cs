using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using ToolVba.App.Services;
using ToolVba.Core;

namespace ToolVba.App.ViewModels;

/// <summary>
/// Single screen (DEC-115): one Excel file and one store product per run. Pressing Create packages the
/// licensing module, filled for that product, into a copy named &lt;productKey&gt;_&lt;version&gt;_ban-quyen.xlsm.
/// </summary>
internal sealed class MainViewModel : ObservableObject
{
    private const string NoProductOutput = "(nhập sản phẩm để xem tên file)";

    private readonly Window _owner;
    private readonly AppServices _services;
    private string? _outputDirectory;
    private WorkbookItem? _workbook;
    private string _status = "Kéo thả một file Excel vào cửa sổ này, nhập sản phẩm rồi bấm Tạo.";
    private bool _hasError;
    private bool _isWorking;
    private int _progressPercent;
    private string? _lastOutput;
    private IReadOnlyList<string> _warnings = [];
    /// <summary>The status line shows the "Xong…" of the last run; it no longer applies once the file or product changes.</summary>
    private bool _showsResult;

    public MainViewModel(Window owner, AppServices services)
    {
        _owner = owner;
        _services = services;
        Package = new VbaPackageStatusViewModel(services.Packages, services.ConfigurationError);
        Update = new AppUpdateViewModel(services.Updater, () => IsWorking);
        Trust = new TrustBannerViewModel(services.Trust, owner);
        Product = new ProductInputViewModel();
        BrowseCommand = new RelayCommand(Browse, () => IsIdle);
        CreateCommand = new RelayCommand(async () => await CreateAsync(), CanCreate);
        ClearCommand = new RelayCommand(() => { Workbook = null; ClearResult(); }, () => IsIdle && HasFile);
        ChangeOutputFolderCommand = new RelayCommand(ChangeOutputFolder, () => IsIdle);
        ResetOutputFolderCommand = new RelayCommand(() => SetOutputDirectory(null), () => IsIdle && _outputDirectory is not null);
        OpenOutputCommand = new RelayCommand(OpenOutput, () => CanOpenOutput);
        OpenFolderCommand = new RelayCommand(OpenFolder, () => CanOpenOutput);
        Product.ProductChanged += (_, _) => { RefreshOutput(); RaiseCommands(); };
        Package.PropertyChanged += (_, _) => RaiseCommands();
        Update.PropertyChanged += (_, e) => { if (e.PropertyName == nameof(AppUpdateViewModel.IsDownloading)) RaiseCommands(); };
    }

    /// <summary>True while work is running that closing the window or restarting would interrupt.</summary>
    public bool IsBusy => IsWorking || Update.IsDownloading;

    public VbaPackageStatusViewModel Package { get; }
    public AppUpdateViewModel Update { get; }
    public TrustBannerViewModel Trust { get; }
    public ProductInputViewModel Product { get; }

    public ICommand BrowseCommand { get; }
    public ICommand CreateCommand { get; }
    public ICommand ClearCommand { get; }
    public ICommand ChangeOutputFolderCommand { get; }
    public ICommand ResetOutputFolderCommand { get; }
    public ICommand OpenOutputCommand { get; }
    public ICommand OpenFolderCommand { get; }

    public string WindowTitle => $"Đóng gói bản quyền — phiên bản {AppConfig.AppVersionText}";
    public bool HasFile => _workbook is not null;
    public string OutputFolderText => _outputDirectory ?? "Lưu cạnh file gốc";
    public bool HasCustomOutputFolder => _outputDirectory is not null;
    public string Status { get => _status; private set => Set(ref _status, value); }
    public bool HasError { get => _hasError; private set => Set(ref _hasError, value); }
    public int ProgressPercent { get => _progressPercent; private set => Set(ref _progressPercent, value); }
    public IReadOnlyList<string> Warnings { get => _warnings; private set => Set(ref _warnings, value); }
    public bool IsIdle => !IsWorking;
    public bool CanOpenOutput => IsIdle && _lastOutput is not null && File.Exists(_lastOutput);

    public WorkbookItem? Workbook
    {
        get => _workbook;
        private set
        {
            if (!Set(ref _workbook, value)) return;
            OnPropertyChanged(nameof(HasFile));
            RaiseCommands();
        }
    }

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

    /// <summary>Takes the first supported file: one file and one product per run.</summary>
    public void AddFiles(IEnumerable<string> paths)
    {
        if (IsWorking) return;
        var all = paths.ToList();
        var supported = all.Where(p => IsSupportedFile(p) && File.Exists(p)).ToList();
        HasError = false;
        if (supported.Count == 0)
        {
            Status = "Chỉ nhận file Excel .xlsx hoặc .xlsm.";
            return;
        }
        Workbook = new WorkbookItem(supported[0], OutputFor(supported[0]), _ => Workbook = null);
        ClearResult();
        Status = all.Count > 1
            ? "Mỗi lần chỉ đóng gói một file cho một sản phẩm; đã lấy file đầu tiên."
            : "Nhập sản phẩm rồi bấm Tạo. File gốc được giữ nguyên.";
    }

    // Blocked while an app update downloads: applying it restarts the app and would cut an install short.
    private bool CanCreate() => IsIdle && !Update.IsDownloading && _workbook is { State: not WorkbookItemState.Done } &&
                                Product.Product is not null && Package.Current is not null;

    private async Task CreateAsync()
    {
        if (_workbook is not { } item || Product.Product is not { } product) return;
        IsWorking = true;
        HasError = false;
        Warnings = [];
        _lastOutput = null;
        try
        {
            var package = await PrepareAsync();
            if (package is null) return;
            await InstallAsync(item, product, package);
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

    private async Task InstallAsync(WorkbookItem item, LicenseProduct product, VbaPackage package)
    {
        item.State = WorkbookItemState.Working;
        item.Message = null;
        item.OutputPath = OutputPathResolver.Resolve(item.SourcePath, product, _outputDirectory);
        var progress = new Progress<VbaInstallProgress>(p =>
        {
            Status = p.Stage;
            ProgressPercent = p.Percent;
        });

        VbaPackageInstallResult result;
        try
        {
            result = await _services.Installer.InstallAsync(new(item.SourcePath, item.OutputPath, package, product), progress);
        }
        finally
        {
            // Never leave the row stuck on "Đang xử lý…" (its remove button stays disabled in that state).
            if (item.State == WorkbookItemState.Working) item.State = WorkbookItemState.Failed;
        }
        item.State = result.Succeeded ? WorkbookItemState.Done : WorkbookItemState.Failed;
        item.Message = result.Succeeded ? null : result.Message;
        ProgressPercent = 100;
        if (!result.Succeeded)
        {
            Trust.Refresh();
            SetFailure(result.Message);
            return;
        }
        _lastOutput = result.OutputPath;
        Warnings = result.Warnings ?? [];
        OnPropertyChanged(nameof(CanOpenOutput));
        Status = $"Xong! Đã tạo {item.OutputName} cho {product.ProductKey} {product.Version} (code VBA v{package.Version}). " +
                 "Mở thử file để kích hoạt rồi mới tải lên cửa hàng. File gốc được giữ nguyên.";
        _showsResult = true;
    }

    private void ClearResult()
    {
        if (!_showsResult) return;
        _showsResult = false;
        Warnings = [];
        Status = "Nhập sản phẩm rồi bấm Tạo. File gốc được giữ nguyên.";
    }

    private string OutputFor(string source) =>
        Product.Product is { } product ? OutputPathResolver.Resolve(source, product, _outputDirectory) : NoProductOutput;

    private void RefreshOutput()
    {
        ClearResult();
        if (_workbook is { State: not WorkbookItemState.Working } item)
        {
            if (item.State == WorkbookItemState.Done) item.State = WorkbookItemState.Ready;
            item.OutputPath = OutputFor(item.SourcePath);
        }
    }

    private void Browse()
    {
        var dialog = new OpenFileDialog
        {
            Title = "Chọn file Excel sản phẩm cần đóng gói bản quyền",
            Filter = "File Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(_owner) == true) AddFiles([dialog.FileName]);
    }

    private void ChangeOutputFolder()
    {
        var dialog = new OpenFolderDialog { Title = "Chọn thư mục lưu file kết quả" };
        if (dialog.ShowDialog(_owner) == true) SetOutputDirectory(dialog.FolderName);
    }

    private void SetOutputDirectory(string? directory)
    {
        _outputDirectory = directory;
        RefreshOutput();
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
