using System.Security.Cryptography;

namespace ToolVba.Core;

public sealed class ExcelModuleInsertionService
{
    private readonly IExcelAutomationHost _excel;
    private readonly ExcelVbaInjector _inspector;
    private readonly BasModuleManifestValidator _moduleValidator;
    private readonly ExcelPackageValidator _packageValidator;
    private readonly SafeOperationAuditLog? _auditLog;

    public ExcelModuleInsertionService(
        IExcelAutomationHost? excel = null,
        ExcelVbaInjector? inspector = null,
        BasModuleManifestValidator? moduleValidator = null,
        ExcelPackageValidator? packageValidator = null,
        SafeOperationAuditLog? auditLog = null)
    {
        _excel = excel ?? new ExcelAutomationHost();
        _inspector = inspector ?? new ExcelVbaInjector();
        _moduleValidator = moduleValidator ?? new BasModuleManifestValidator();
        _packageValidator = packageValidator ?? new ExcelPackageValidator();
        _auditLog = auditLog;
    }

    public Task<ExcelModuleInsertionResult> InsertAsync(
        ExcelModuleInsertionRequest request,
        IProgress<ExcelModuleOperationProgress>? progress = null,
        CancellationToken cancellationToken = default) =>
        StaTask.Run(() =>
        {
            var result = Insert(request, progress, cancellationToken);
            if (_auditLog is not null && !_auditLog.TryWrite(request, result))
                result = result with { Warnings = (result.Warnings ?? []).Append("Không ghi được log metadata cục bộ.").ToArray() };
            return result;
        }, cancellationToken);

    private ExcelModuleInsertionResult Insert(
        ExcelModuleInsertionRequest request,
        IProgress<ExcelModuleOperationProgress>? progress,
        CancellationToken cancellationToken)
    {
        string input;
        string module;
        string manifest;
        string output;
        try
        {
            input = Path.GetFullPath(request.InputWorkbookPath);
            module = Path.GetFullPath(request.ModulePath);
            manifest = Path.GetFullPath(request.ManifestPath);
            output = Path.GetFullPath(request.OutputPath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return ExcelModuleInsertionResult.Failure("Có đường dẫn không hợp lệ hoặc quá dài.");
        }

        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase))
            return ExcelModuleInsertionResult.Failure("File kết quả không được trùng đường dẫn file nguồn.");
        var isXlsx = Path.GetExtension(input).Equals(".xlsx", StringComparison.OrdinalIgnoreCase);
        if (!isXlsx && !Path.GetExtension(input).Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            return ExcelModuleInsertionResult.Failure("Chế độ thêm module chỉ hỗ trợ workbook nguồn .xlsx hoặc .xlsm.");
        if (!Path.GetExtension(output).Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            return ExcelModuleInsertionResult.Failure("File kết quả phải có phần mở rộng .xlsm.");
        if (!File.Exists(input))
            return ExcelModuleInsertionResult.Failure("Không tìm thấy workbook nguồn.");
        if (!Directory.Exists(Path.GetDirectoryName(output)))
            return ExcelModuleInsertionResult.Failure("Thư mục lưu kết quả không tồn tại.");
        if (File.Exists(output) && !request.ConfirmOutputOverwrite)
            return ExcelModuleInsertionResult.Failure("File kết quả đã tồn tại. Cần xác nhận ghi đè trước khi tiếp tục.");

        var sourceInspection = _inspector.InspectWorkbook(input);
        if (!sourceInspection.IsValid)
            return ExcelModuleInsertionResult.Failure(sourceInspection.Error ?? "Workbook nguồn không hợp lệ.");

        var moduleCheck = _moduleValidator.Validate(module, manifest);
        if (!moduleCheck.IsValid || moduleCheck.ModuleName is null)
            return ExcelModuleInsertionResult.Failure(moduleCheck.Error ?? "Module không vượt qua kiểm tra manifest.");

        var warnings = new List<string>();
        if (sourceInspection.HasSignatureArtifacts)
            warnings.Add("Nguồn có artifact chữ ký package hoặc VBA. Thay đổi project có thể làm chữ ký cũ mất hiệu lực; công cụ không xác minh chữ ký.");
        warnings.Add("Macro bị vô hiệu trong phiên đóng gói và không được chạy. Kiểm tra output trong Excel trên máy thử nghiệm theo chính sách tổ chức.");

        var tempBase = Path.Combine(Path.GetDirectoryName(output)!, $".tool-vba-{Guid.NewGuid():N}.tmp");
        var tempPath = tempBase + (isXlsx ? ".xlsx" : ".xlsm");
        var convertedPath = isXlsx ? tempBase + ".xlsm" : tempPath;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new("Đang ghi nhận hash nguồn và tạo bản tạm.", 10));
            var inputHash = HashFile(input);
            File.Copy(input, tempPath, overwrite: false);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new("Đang mở bản tạm trong Excel với macro bị vô hiệu.", 30));
            IReadOnlyList<string> before;
            using (var session = _excel.Open(tempPath, readOnly: false))
            {
                before = session.GetComponentNames();
                if (before.Any(name => string.Equals(name, moduleCheck.ModuleName, StringComparison.OrdinalIgnoreCase)))
                    return ExcelModuleInsertionResult.Failure("Tên module đã tồn tại trong VBA Project. Không có module cũ nào bị thay đổi.");

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(new("Đang import module đã khớp manifest.", 55));
                var importedName = session.ImportStandardModule(module);
                if (!string.Equals(importedName, moduleCheck.ModuleName, StringComparison.OrdinalIgnoreCase))
                    return ExcelModuleInsertionResult.Failure("Tên component Excel import không khớp manifest. Bản tạm sẽ bị loại bỏ.");

                var afterImport = session.GetComponentNames();
                if (!PreservesComponents(before, afterImport, moduleCheck.ModuleName, sourceInspection.HasVbaProject))
                    return ExcelModuleInsertionResult.Failure("Không xác nhận được module vừa import hoặc component VBA gốc trong Excel. " +
                        DescribeComponentCheck(before, afterImport, moduleCheck.ModuleName, sourceInspection.HasVbaProject));
                if (isXlsx)
                    session.SaveAsMacroEnabled(convertedPath);
                else
                    session.Save();
            }

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new("Đang mở lại bản tạm để xác nhận component đã lưu.", 70));
            using (var verificationSession = _excel.Open(convertedPath, readOnly: true))
            {
                var savedComponents = verificationSession.GetComponentNames();
                if (!PreservesComponents(before, savedComponents, moduleCheck.ModuleName, sourceInspection.HasVbaProject))
                    return ExcelModuleInsertionResult.Failure("Output tạm không giữ đủ VBA sau khi lưu. " +
                        DescribeComponentCheck(before, savedComponents, moduleCheck.ModuleName, sourceInspection.HasVbaProject));
            }

            progress?.Report(new("Đang kiểm tra package Open XML và hash module.", 85));
            var packageResult = _packageValidator.ValidateMacroWorkbook(convertedPath);
            if (!packageResult.IsValid)
                return ExcelModuleInsertionResult.Failure("Kiểm tra package thất bại: " + string.Join(" ", packageResult.Errors));

            var moduleCheckAfter = _moduleValidator.Validate(module, manifest);
            if (!moduleCheckAfter.IsValid || !string.Equals(moduleCheck.Sha256, moduleCheckAfter.Sha256, StringComparison.OrdinalIgnoreCase))
                return ExcelModuleInsertionResult.Failure("Module hoặc manifest đã thay đổi trong quá trình xử lý.");
            if (!CryptographicOperations.FixedTimeEquals(inputHash, HashFile(input)))
                return ExcelModuleInsertionResult.Failure("Hash file nguồn thay đổi trong lúc xử lý. Không xuất output.");

            cancellationToken.ThrowIfCancellationRequested();
            if (File.Exists(output) && !request.ConfirmOutputOverwrite)
                return ExcelModuleInsertionResult.Failure("File kết quả vừa xuất hiện và chưa được xác nhận ghi đè.");

            progress?.Report(new("Đang hoàn tất output.", 95));
            File.Move(convertedPath, output, overwrite: request.ConfirmOutputOverwrite);
            progress?.Report(new("Hoàn tất.", 100));
            return ExcelModuleInsertionResult.Success(output, warnings);
        }
        catch (OperationCanceledException)
        {
            return ExcelModuleInsertionResult.Failure("Đã hủy thao tác. File nguồn và output có sẵn được giữ nguyên.");
        }
        catch (ExcelAutomationException ex)
        {
            return ExcelModuleInsertionResult.Failure(ex.Message);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            return ExcelModuleInsertionResult.Failure("Không thể tạo output. Kiểm tra quyền truy cập Excel/file và dung lượng ổ đĩa.");
        }
        catch (Exception)
        {
            return ExcelModuleInsertionResult.Failure("Excel không hoàn tất được thao tác. File nguồn được giữ nguyên; kiểm tra cấu hình Excel và thử trên bản sao.");
        }
        finally
        {
            try
            {
                if (File.Exists(tempPath)) File.Delete(tempPath);
                if (isXlsx && File.Exists(convertedPath)) File.Delete(convertedPath);
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static bool PreservesComponents(IReadOnlyList<string> before, IReadOnlyList<string> after,
        string newModule, bool sourceHadVbaProject) =>
        after.Distinct(StringComparer.OrdinalIgnoreCase).Count() == after.Count &&
        after.Any(name => string.Equals(name, newModule, StringComparison.OrdinalIgnoreCase)) &&
        (!sourceHadVbaProject || before.All(oldName => after.Any(name => string.Equals(name, oldName, StringComparison.OrdinalIgnoreCase))));

    private static string DescribeComponentCheck(IReadOnlyList<string> before, IReadOnlyList<string> after,
        string newModule, bool sourceHadVbaProject)
    {
        var modulePresent = after.Any(name => string.Equals(name, newModule, StringComparison.OrdinalIgnoreCase));
        var missingOriginal = sourceHadVbaProject
            ? before.Count(oldName => !after.Any(name => string.Equals(name, oldName, StringComparison.OrdinalIgnoreCase)))
            : 0;
        return $"Trước: {before.Count} component; sau: {after.Count}; module mới: {(modulePresent ? "có" : "thiếu")}; " +
               $"component VBA gốc thiếu: {missingOriginal}. Bản tạm đã bị loại bỏ; file nguồn không bị sửa.";
    }

    private static byte[] HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        return SHA256.HashData(stream);
    }
}
