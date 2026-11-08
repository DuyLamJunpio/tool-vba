using System.Text;
using System.Text.RegularExpressions;

namespace ToolVba.Core;

/// <param name="Module">The licensing module of the package.</param>
/// <param name="ConfiguredPath">A copy of it with this product's settings, imported instead of the package file.</param>
internal sealed record LicenseSetup(VbaPackageComponentFile Module, string ConfiguredPath, SheetPlan Sheets, LicenseProduct Product);

/// <summary>
/// The licensing part of packaging a product (DEC-112, DEC-115): find the licensing module of the package,
/// fill it for the product, show only the guide sheet and protect every other sheet.
/// </summary>
internal static class LicensePackaging
{
    private const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Multiline;
    private static readonly Regex DefinesKey = new(@"^[ \t]*(?:Public[ \t]+)?Function[ \t]+Cfg_ProductKey\b", Options);
    private static readonly Regex AutoOpen = new(@"^[ \t]*(?:(?:Public|Private)[ \t]+)?Sub[ \t]+Auto_Open\b", Options);

    public static (VbaPackageComponentFile? Module, string? Error) FindLicenseModule(VbaPackage package)
    {
        var modules = package.CodeModules
            .Where(c => c.Kind == VbaComponentKind.Standard && DefinesKey.IsMatch(File.ReadAllText(c.FilePath, Encoding.Latin1)))
            .ToList();
        return modules.Count switch
        {
            1 => (modules[0], null),
            0 => (null, "Gói code VBA đang dùng chưa có module bản quyền (modLicenseApiV1). " +
                        "Hãy xuất bản lại gói từ file mẫu bản quyền trên Admin web."),
            _ => (null, "Gói code VBA có nhiều module định nghĩa Cfg_ProductKey. Hãy xuất bản lại gói từ file mẫu bản quyền.")
        };
    }

    /// <summary>The product's own standard modules must not define Auto_Open: the licensing module's Auto_Open is the gate.</summary>
    public static string? FindAutoOpenConflict(
        IVbaWorkbookSession session, IReadOnlyList<VbaComponentInfo> components, IReadOnlyCollection<string> replaced)
    {
        var replacedNames = replaced.ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var module in components.Where(c => c.Kind == VbaComponentKind.Standard && !replacedNames.Contains(c.Name)))
        {
            if (AutoOpen.IsMatch(session.GetDocumentCode(module.Name)))
                return $"Module {module.Name} của file đã có thủ tục Auto_Open. Module bản quyền dùng Auto_Open để kiểm tra bản quyền khi mở file; " +
                       "hãy chuyển code đó sang Workbook_Open hoặc đổi tên thủ tục rồi thử lại.";
        }
        return null;
    }

    public static (LicenseSetup? Setup, string? Error) Prepare(
        IVbaWorkbookSession session, VbaPackageComponentFile module, LicenseProduct product, string directory)
    {
        var workbookSheets = session.GetSheets();
        var (sheets, sheetError) = GuideSheet.Plan(workbookSheets, session.IsStructureProtected(), IsToolGuide(session, workbookSheets));
        if (sheets is null) return (null, sheetError);
        var (code, configError) = LicenseModuleConfigurator.Configure(
            File.ReadAllText(module.FilePath, Encoding.Latin1), product, sheets.ProtectedNames);
        if (code is null) return (null, configError);
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, Path.GetFileName(module.FilePath));
        File.WriteAllBytes(path, Encoding.Latin1.GetBytes(code));
        return (new LicenseSetup(module, path, sheets, product), null);
    }

    /// <summary>
    /// The marker property only names the sheet; a sheet the author made later under the same name would be
    /// wiped when reused, so its first line must also be the tool's own text.
    /// </summary>
    private static bool IsToolGuide(IVbaWorkbookSession session, IReadOnlyList<WorkbookSheetInfo> sheets)
    {
        var guide = sheets.FirstOrDefault(s => string.Equals(s.Name, GuideSheet.Name, StringComparison.OrdinalIgnoreCase));
        return guide is not null &&
               string.Equals(session.GetCustomProperty(GuideSheet.MarkerProperty), GuideSheet.Name, StringComparison.Ordinal) &&
               string.Equals(session.ReadCell(guide.Name, "A1"), GuideSheet.Lines[0], StringComparison.Ordinal);
    }

    /// <summary>The guide sheet first and visible, then every other sheet very hidden (Excel needs one visible sheet).</summary>
    public static void ApplySheets(IVbaWorkbookSession session, LicenseSetup setup)
    {
        session.WriteGuideSheet(GuideSheet.Name, GuideSheet.Lines, setup.Sheets.CreateGuide);
        foreach (var name in setup.Sheets.ProtectedNames)
            session.SetSheetVisibility(name, SheetVisibility.VeryHidden);
        session.SetCustomProperty(GuideSheet.MarkerProperty, GuideSheet.Name);
    }

    /// <summary>Checks the saved copy: product settings in the module, only the guide sheet visible. Null when all is right.</summary>
    public static string? VerifySaved(IVbaWorkbookSession session, LicenseSetup setup)
    {
        var code = session.GetDocumentCode(setup.Module.Name);
        if (!code.Contains($"Cfg_ProductKey = \"{setup.Product.ProductKey}\"", StringComparison.Ordinal) ||
            !code.Contains($"Cfg_ProductVersion = \"{setup.Product.Version}\"", StringComparison.Ordinal))
            return "Module bản quyền trong file tạm chưa có mã sản phẩm và phiên bản.";
        var sheets = session.GetSheets();
        var guide = sheets.FirstOrDefault(s => string.Equals(s.Name, GuideSheet.Name, StringComparison.OrdinalIgnoreCase));
        if (guide is null || guide.Visibility != SheetVisibility.Visible)
            return "Sheet hướng dẫn chưa được lưu hoặc đang bị ẩn.";
        var exposed = sheets.FirstOrDefault(s => !ReferenceEquals(s, guide) && s.Visibility != SheetVisibility.VeryHidden);
        return exposed is null ? null : $"Sheet \"{exposed.Name}\" chưa được ẩn trong file tạm.";
    }
}
