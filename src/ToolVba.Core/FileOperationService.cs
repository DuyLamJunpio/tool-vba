namespace ToolVba.Core;

public sealed class FileOperationService
{
    private readonly ExcelVbaInjector _injector;
    private readonly ExcelPackageValidator _validator;

    public FileOperationService(ExcelVbaInjector? injector = null, ExcelPackageValidator? validator = null)
    {
        _injector = injector ?? new ExcelVbaInjector();
        _validator = validator ?? new ExcelPackageValidator();
    }

    public Task<OperationResult> CreateAsync(
        InjectionRequest request,
        IProgress<OperationProgress>? progress = null,
        CancellationToken cancellationToken = default) =>
        Task.Run(() => Create(request, progress, cancellationToken), cancellationToken);

    private OperationResult Create(InjectionRequest request, IProgress<OperationProgress>? progress, CancellationToken cancellationToken)
    {
        string input;
        string vba;
        string output;
        try
        {
            input = Path.GetFullPath(request.InputWorkbookPath);
            vba = Path.GetFullPath(request.VbaProjectPath);
            output = Path.GetFullPath(request.OutputPath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return OperationResult.Failure("Đường dẫn file không hợp lệ hoặc quá dài.");
        }

        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase))
            return OperationResult.Failure("Nơi lưu file kết quả không được trùng với file Excel nguồn.");
        if (!Path.GetExtension(input).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !Path.GetExtension(input).Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            return OperationResult.Failure("Chỉ hỗ trợ file Excel nguồn .xlsx hoặc .xlsm.");
        if (!Path.GetExtension(vba).Equals(".bin", StringComparison.OrdinalIgnoreCase))
            return OperationResult.Failure("VBA project phải là file .bin.");
        if (!Path.GetExtension(output).Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            return OperationResult.Failure("File kết quả phải có phần mở rộng .xlsm.");
        if (!File.Exists(vba))
            return OperationResult.Failure("Không tìm thấy file VBA project đã chọn.");
        try
        {
            if (new FileInfo(vba).Length == 0)
                return OperationResult.Failure("File VBA project đang trống.");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return OperationResult.Failure("Không thể đọc file VBA project. Kiểm tra quyền truy cập file.");
        }

        var outputDirectory = Path.GetDirectoryName(output);
        if (string.IsNullOrEmpty(outputDirectory) || !Directory.Exists(outputDirectory))
            return OperationResult.Failure("Thư mục lưu kết quả không tồn tại hoặc không truy cập được.");
        if (File.Exists(output) && !request.ConfirmOutputOverwrite)
            return OperationResult.Failure("File kết quả đã tồn tại. Xác nhận ghi đè trước khi tiếp tục.");

        var inspection = _injector.InspectWorkbook(input);
        if (!inspection.IsValid)
            return OperationResult.Failure(inspection.Error ?? "File Excel nguồn không hợp lệ.");
        if (inspection.HasVbaProject && !request.ConfirmVbaReplacement)
            return OperationResult.Failure("Workbook đã có VBA project. Chọn “Thay VBA project hiện có” và xác nhận cảnh báo để tiếp tục.");

        var warnings = new List<string>();
        if (inspection.HasVbaProject)
            warnings.Add("VBA project hiện có sẽ được thay thế theo xác nhận của bạn.");
        if (inspection.HasSignatureArtifacts)
            warnings.Add("Nguồn có thành phần chữ ký. Việc sửa package hoặc thay VBA có thể làm chữ ký cũ mất hiệu lực; công cụ không xác minh trạng thái chữ ký.");
        warnings.Add("Các kiểm tra cấu trúc không xác nhận VBA an toàn, hợp lệ nghiệp vụ hay chạy đúng. Hãy mở bản kết quả trong Excel trên máy kiểm thử và kiểm tra theo chính sách bảo mật của tổ chức.");

        var tempPath = Path.Combine(outputDirectory, $".tool-vba-{Guid.NewGuid():N}.tmp.xlsm");
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new("Đang tạo bản tạm, file nguồn sẽ không bị sửa.", 10));
            File.Copy(input, tempPath, overwrite: false);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new("Đang chèn VBA project vào bản tạm.", 45));
            var injectResult = _injector.Inject(tempPath, vba, request.ConfirmVbaReplacement);
            if (!injectResult.Succeeded)
                return OperationResult.Failure(injectResult.Message);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new("Đang kiểm tra package, relationship và SHA-256.", 75));
            var validation = _validator.Validate(tempPath, vba);
            if (!validation.IsValid)
                return OperationResult.Failure("Kiểm tra đầu ra thất bại: " + string.Join(" ", validation.Errors));

            cancellationToken.ThrowIfCancellationRequested();
            if (File.Exists(output) && !request.ConfirmOutputOverwrite)
                return OperationResult.Failure("File kết quả vừa xuất hiện và chưa được xác nhận ghi đè.");

            progress?.Report(new("Đang hoàn tất file kết quả.", 95));
            File.Move(tempPath, output, overwrite: request.ConfirmOutputOverwrite);
            progress?.Report(new("Hoàn tất.", 100));
            return OperationResult.Success(output, warnings);
        }
        catch (OperationCanceledException)
        {
            return OperationResult.Failure("Đã hủy thao tác. File nguồn và file kết quả cũ được giữ nguyên.");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            return OperationResult.Failure("Không thể tạo file kết quả. Kiểm tra quyền ghi, dung lượng ổ đĩa và đường dẫn thư mục.");
        }
        finally
        {
            try
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
            catch (IOException)
            {
                // Best-effort cleanup; no document content is logged.
            }
            catch (UnauthorizedAccessException)
            {
                // Best-effort cleanup; no document content is logged.
            }
        }
    }
}
