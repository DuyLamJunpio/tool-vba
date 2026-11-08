namespace ToolVba.Core;

/// <param name="Product">The store product and version the licensing module is filled for (DEC-115).</param>
public sealed record VbaPackageInstallRequest(
    string InputWorkbookPath,
    string OutputPath,
    VbaPackage Package,
    LicenseProduct Product,
    bool ConfirmOutputOverwrite = false);

public sealed record VbaPackageInstallResult(
    bool Succeeded,
    string Message,
    string? OutputPath = null,
    IReadOnlyList<string>? Warnings = null)
{
    public static VbaPackageInstallResult Success(string output, string version, IReadOnlyList<string> warnings) =>
        new(true, $"Đã dán code VBA v{version} và kiểm tra file kết quả.", output, warnings);

    public static VbaPackageInstallResult Failure(string message) => new(false, message);
}

public sealed record VbaInstallProgress(string Stage, int Percent);

internal sealed record DocumentTarget(VbaPackageComponentFile Source, string TargetName);

/// <summary>
/// Decides, before anything is changed, which components are replaced, which are kept and where
/// document-module code goes. Components the package (or an earlier package) manages are replaced;
/// everything else in the workbook is kept.
/// </summary>
internal sealed record VbaInstallPlan(
    IReadOnlyList<string> ComponentsToRemove,
    IReadOnlyList<DocumentTarget> Documents,
    IReadOnlyList<string> PreservedNames)
{
    public static (VbaInstallPlan? Plan, string? Error) Create(
        IReadOnlyList<VbaComponentInfo> existing,
        IReadOnlyList<string> previouslyManaged,
        VbaPackage package,
        bool sourceHadVbaProject)
    {
        if (package.DocumentModules.Any(d => d.Kind == VbaComponentKind.Workbook) &&
            existing.Any(c => c.Kind == VbaComponentKind.Workbook && c.Name.Length == 0))
            return (null, "Module ThisWorkbook của file này bị Excel làm mất tên nên không thể dán code vào. " +
                          "Hãy mở file trong Excel, Alt+F11, đặt lại tên (Name) cho ThisWorkbook rồi lưu và thử lại.");

        var (documents, missing) = MapDocuments(existing, package);
        if (missing.Count > 0)
            return (null, "File Excel này không có sheet mà code công ty cần: " + string.Join(", ", missing) +
                          ". Code của sheet được dán theo CodeName giống file mẫu.");

        var packageNames = package.CodeModules.Select(c => c.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var clash = existing.FirstOrDefault(c => !IsCodeModule(c.Kind) && packageNames.Contains(c.Name));
        if (clash is not null)
            return (null, $"Module {clash.Name} của code công ty trùng tên với một sheet hoặc ThisWorkbook trong file Excel.");

        var managed = packageNames.Union(previouslyManaged, StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var remove = existing.Where(c => IsCodeModule(c.Kind) && managed.Contains(c.Name)).Select(c => c.Name).ToList();
        var preserved = sourceHadVbaProject
            ? existing.Select(c => c.Name).Where(n => !remove.Contains(n, StringComparer.OrdinalIgnoreCase)).ToList()
            : [];
        return (new(remove, documents, preserved), null);
    }

    public static bool IsCodeModule(VbaComponentKind kind) =>
        kind is VbaComponentKind.Standard or VbaComponentKind.Class or VbaComponentKind.Form;

    private static (List<DocumentTarget> Documents, List<string> Missing) MapDocuments(
        IReadOnlyList<VbaComponentInfo> existing, VbaPackage package)
    {
        var documents = new List<DocumentTarget>();
        var missing = new List<string>();
        var workbook = existing.FirstOrDefault(c => c.Kind == VbaComponentKind.Workbook);
        foreach (var document in package.DocumentModules)
        {
            var target = document.Kind == VbaComponentKind.Workbook ? workbook : FindSheet(existing, document);
            if (target is null) missing.Add(document.SheetName is null ? document.Name : $"{document.Name} (sheet \"{document.SheetName}\")");
            else documents.Add(new(document, target.Name));
        }
        return (documents, missing);
    }

    /// <summary>Matches by CodeName first; Excel in another language may name it differently, so fall back to the tab name.</summary>
    private static VbaComponentInfo? FindSheet(IReadOnlyList<VbaComponentInfo> existing, VbaPackageComponentFile document)
    {
        var sheets = existing.Where(c => c.Kind == VbaComponentKind.Sheet).ToList();
        return sheets.FirstOrDefault(c => string.Equals(c.Name, document.Name, StringComparison.OrdinalIgnoreCase))
               ?? (document.SheetName is null
                   ? null
                   : sheets.FirstOrDefault(c => string.Equals(c.SheetName, document.SheetName, StringComparison.OrdinalIgnoreCase)));
    }
}
