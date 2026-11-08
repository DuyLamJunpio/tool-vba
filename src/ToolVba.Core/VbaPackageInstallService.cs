using System.Security.Cryptography;
using System.Text;

namespace ToolVba.Core;

/// <summary>
/// Packages a product: pastes the company VBA package, with its licensing module filled for the chosen
/// store product, into a copy of a workbook, adds the guide sheet and protects every other sheet
/// (DEC-112, DEC-115). The source file is never modified: Excel works on a temporary copy with macros
/// disabled, the saved copy is reopened and checked, and only then is it moved to the output path.
/// </summary>
public sealed class VbaPackageInstallService(
    IVbaWorkbookHost? excel = null,
    ExcelVbaInjector? inspector = null,
    ExcelPackageValidator? packageValidator = null,
    VbaPackageReader? reader = null,
    SafeOperationAuditLog? auditLog = null)
{
    private readonly IVbaWorkbookHost _excel = excel ?? new ExcelAutomationHost();
    private readonly ExcelVbaInjector _inspector = inspector ?? new ExcelVbaInjector();
    private readonly ExcelPackageValidator _packageValidator = packageValidator ?? new ExcelPackageValidator();
    private readonly VbaPackageReader _reader = reader ?? new VbaPackageReader();

    public Task<VbaPackageInstallResult> InstallAsync(
        VbaPackageInstallRequest request,
        IProgress<VbaInstallProgress>? progress = null,
        CancellationToken cancellationToken = default) =>
        StaTask.Run(() =>
        {
            var result = Install(request, progress, cancellationToken);
            if (auditLog is not null && !auditLog.TryWritePackageInstall(request, result))
                result = result with { Warnings = [.. result.Warnings ?? [], "Không ghi được log metadata cục bộ."] };
            return result;
        }, cancellationToken);

    private VbaPackageInstallResult Install(VbaPackageInstallRequest request, IProgress<VbaInstallProgress>? progress, CancellationToken ct)
    {
        var (paths, pathError) = InstallPaths.Resolve(request);
        if (paths is null) return VbaPackageInstallResult.Failure(pathError!);
        var inspection = _inspector.InspectWorkbook(paths.Input);
        if (!inspection.IsValid) return VbaPackageInstallResult.Failure(inspection.Error ?? "Workbook nguồn không hợp lệ.");
        var packageCheck = _reader.Verify(request.Package);
        if (!packageCheck.IsValid) return VbaPackageInstallResult.Failure(packageCheck.Error!);

        IReadOnlyList<string> warnings = inspection.HasSignatureArtifacts
            ? ["File nguồn có chữ ký số. Sau khi dán code, chữ ký cũ có thể mất hiệu lực."]
            : [];
        using var work = new InstallWorkArea(paths);
        try
        {
            progress?.Report(new("Đang chuẩn bị bản tạm và code VBA.", 10));
            var package = work.StagePackage(request.Package, _reader);
            if (package is null) return VbaPackageInstallResult.Failure("Gói code VBA thay đổi trong lúc chuẩn bị. Hãy thử lại.");
            var (licenseModule, licenseError) = LicensePackaging.FindLicenseModule(package);
            if (licenseModule is null) return VbaPackageInstallResult.Failure(licenseError!);
            var inputHash = HashFile(paths.Input);
            File.Copy(paths.Input, work.TempPath, overwrite: false);

            ct.ThrowIfCancellationRequested();
            progress?.Report(new("Đang mở bản tạm trong Excel với macro bị vô hiệu.", 30));
            var target = new ApplyTarget(work, package, licenseModule, request.Product, inspection.HasVbaProject);
            var (outcome, applyError) = Apply(target, progress, ct);
            if (outcome is null) return VbaPackageInstallResult.Failure(applyError! + " File nguồn không bị sửa.");

            ct.ThrowIfCancellationRequested();
            progress?.Report(new("Đang mở lại bản tạm để xác nhận code VBA và sheet đã lưu.", 70));
            var verifyError = VerifySaved(work.SavedPath, package, outcome);
            if (verifyError is not null) return VbaPackageInstallResult.Failure(verifyError + " Bản tạm đã bị loại bỏ; file nguồn không bị sửa.");

            progress?.Report(new("Đang kiểm tra package Open XML.", 85));
            var finalError = FinalCheck(work.SavedPath, package, paths, inputHash, request.ConfirmOutputOverwrite);
            if (finalError is not null) return VbaPackageInstallResult.Failure(finalError);

            progress?.Report(new("Đang hoàn tất file kết quả.", 95));
            File.Move(work.SavedPath, paths.Output, overwrite: request.ConfirmOutputOverwrite);
            progress?.Report(new("Hoàn tất.", 100));
            return VbaPackageInstallResult.Success(paths.Output, package.Version, warnings);
        }
        catch (OperationCanceledException)
        {
            return VbaPackageInstallResult.Failure("Đã hủy thao tác. File nguồn và file kết quả có sẵn được giữ nguyên.");
        }
        catch (ExcelAutomationException ex)
        {
            return VbaPackageInstallResult.Failure(ex.Message);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            return VbaPackageInstallResult.Failure($"Không thể tạo file kết quả. Kiểm tra quyền truy cập thư mục, Excel và dung lượng ổ đĩa. (mã: {ex.GetType().Name})");
        }
        catch (Exception ex)
        {
            // Only the exception type is reported: messages from Excel can contain file paths or code.
            return VbaPackageInstallResult.Failure($"Excel không hoàn tất được thao tác. File nguồn được giữ nguyên; hãy đóng Excel và thử lại. (mã: {ex.GetType().Name})");
        }
    }

    private const string WorkbookModuleName = "ThisWorkbook";

    private sealed record ApplyTarget(
        InstallWorkArea Work, VbaPackage Package, VbaPackageComponentFile LicenseModule, LicenseProduct Product, bool SourceHadVbaProject);

    private sealed record ApplyOutcome(VbaInstallPlan Plan, IReadOnlyList<string> WrittenDocuments, LicenseSetup License);

    private (ApplyOutcome? Outcome, string? Error) Apply(ApplyTarget target, IProgress<VbaInstallProgress>? progress, CancellationToken ct)
    {
        using var session = _excel.Open(target.Work.TempPath, readOnly: false);
        var components = NormalizeWorkbookModuleName(session);
        var (plan, planError) = VbaInstallPlan.Create(components, ManagedComponentList.Read(session), target.Package, target.SourceHadVbaProject);
        if (plan is null) return (null, planError);
        var foreign = LicensePackaging.FindForeignLicensingCode(session, components, plan.ComponentsToRemove);
        if (foreign is not null) return (null, foreign);
        var (license, licenseError) = LicensePackaging.Prepare(session, target.LicenseModule, target.Product, target.Work.LicenseDirectory);
        if (license is null) return (null, licenseError);
        var (merges, mergeError) = MergeDocuments(session, plan);
        if (merges is null) return (null, mergeError);

        ct.ThrowIfCancellationRequested();
        progress?.Report(new("Đang thay module cũ bằng code mới.", 45));
        foreach (var name in plan.ComponentsToRemove)
            session.RemoveComponent(name);
        foreach (var component in target.Package.CodeModules)
        {
            var file = string.Equals(component.Name, license.Module.Name, StringComparison.OrdinalIgnoreCase)
                ? license.ConfiguredPath
                : component.FilePath;
            var imported = session.ImportComponent(file);
            if (!string.Equals(imported, component.Name, StringComparison.OrdinalIgnoreCase))
                return (null, $"Excel đặt tên '{imported}' cho module {component.Name} (có thể do trùng tên).");
        }

        progress?.Report(new("Đang dán code vào ThisWorkbook, thêm sheet hướng dẫn và ẩn các sheet khác.", 55));
        foreach (var (document, code) in merges)
            session.SetDocumentCode(document, code);
        ManagedComponentList.Write(session, target.Package.CodeModules.Select(c => c.Name), target.Package.Version);
        LicensePackaging.ApplySheets(session, license);

        var missing = MissingComponents(session.GetComponents(), target.Package, plan);
        if (missing.Count > 0) return (null, "Không xác nhận được code VBA sau khi dán. Thiếu: " + string.Join(", ", missing) + ".");
        if (target.Work.ConvertsToMacroEnabled) session.SaveAsMacroEnabled(target.Work.SavedPath);
        else session.Save();
        return (new(plan, merges.Select(m => m.Target).ToList(), license), null);
    }

    /// <summary>
    /// Excel in Vietnamese names the workbook module "SổlàmviệcNày". Characters outside the ANSI code page
    /// are lost when Excel saves the VBA project, leaving a module with an empty name whose events no longer
    /// fire and which crashes Excel when read. Renaming it to ThisWorkbook first avoids that; VBA's
    /// ThisWorkbook keyword does not depend on the module name.
    /// </summary>
    private static IReadOnlyList<VbaComponentInfo> NormalizeWorkbookModuleName(IVbaWorkbookSession session)
    {
        var components = session.GetComponents();
        var workbook = components.FirstOrDefault(c => c.Kind == VbaComponentKind.Workbook);
        if (workbook is null || workbook.Name.Length == 0 || workbook.Name.All(char.IsAscii) ||
            components.Any(c => string.Equals(c.Name, WorkbookModuleName, StringComparison.OrdinalIgnoreCase)))
            return components;
        session.RenameDocumentModule(workbook.Name, WorkbookModuleName);
        return session.GetComponents();
    }

    private static (List<(string Target, string Code)>? Merges, string? Error) MergeDocuments(IVbaWorkbookSession session, VbaInstallPlan plan)
    {
        var merges = new List<(string, string)>();
        foreach (var document in plan.Documents)
        {
            var packageCode = File.ReadAllText(document.Source.FilePath, Encoding.UTF8);
            if (DocumentModuleMerger.IsEffectivelyEmpty(packageCode)) continue;
            var merged = DocumentModuleMerger.Merge(session.GetDocumentCode(document.TargetName), packageCode);
            if (!merged.Succeeded)
                // Renaming or removing is the remedy for a duplicate name only; other errors carry their own advice.
                return (null, merged.ConflictingNames.Count > 0
                    ? $"{document.TargetName}: {merged.Error} Hãy đổi tên hoặc xóa phần code trùng trong file Excel rồi thử lại."
                    : $"{document.TargetName}: {merged.Error}");
            merges.Add((document.TargetName, merged.Code!));
        }
        return (merges, null);
    }

    private string? VerifySaved(string savedPath, VbaPackage package, ApplyOutcome outcome)
    {
        using var session = _excel.Open(savedPath, readOnly: true);
        var components = session.GetComponents();
        var missing = MissingComponents(components, package, outcome.Plan);
        if (missing.Count > 0) return "File tạm không giữ đủ VBA sau khi lưu. Thiếu: " + string.Join(", ", missing) + ".";
        // Only the modules written above are read back; other document modules may be unreadable.
        foreach (var target in outcome.WrittenDocuments)
        {
            if (!components.Any(c => string.Equals(c.Name, target, StringComparison.OrdinalIgnoreCase)) ||
                !DocumentModuleMerger.ContainsManagedCode(session.GetDocumentCode(target)))
                return $"Code của {target} chưa được lưu.";
        }
        if (session.GetCustomProperty(ManagedComponentList.VersionProperty) != package.Version)
            return "Không ghi được phiên bản code VBA vào file.";
        return LicensePackaging.VerifySaved(session, outcome.License);
    }

    private static List<string> MissingComponents(IReadOnlyList<VbaComponentInfo> components, VbaPackage package, VbaInstallPlan plan)
    {
        var present = components.Select(c => c.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        return package.CodeModules.Select(c => c.Name).Concat(plan.PreservedNames)
            .Where(name => !present.Contains(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>Last checks before the temporary copy becomes the output; null when everything passed.</summary>
    private string? FinalCheck(string savedPath, VbaPackage package, InstallPaths paths, byte[] inputHash, bool confirmOverwrite)
    {
        var packageResult = _packageValidator.ValidateMacroWorkbook(savedPath);
        if (!packageResult.IsValid) return "Kiểm tra package thất bại: " + string.Join(" ", packageResult.Errors);
        if (!_reader.Verify(package).IsValid) return "Gói code VBA thay đổi trong quá trình xử lý. Không xuất file kết quả.";
        if (!CryptographicOperations.FixedTimeEquals(inputHash, HashFile(paths.Input)))
            return "File nguồn thay đổi trong lúc xử lý. Không xuất file kết quả.";
        return File.Exists(paths.Output) && !confirmOverwrite ? "File kết quả vừa xuất hiện và chưa được xác nhận ghi đè." : null;
    }

    private static byte[] HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        return SHA256.HashData(stream);
    }
}
