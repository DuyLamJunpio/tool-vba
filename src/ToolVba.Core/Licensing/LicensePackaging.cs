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
    private static readonly Regex AutoOpen = new(@"^[ \t]*(?:(?:Public|Private|Friend)[ \t]+)?(?:Static[ \t]+)?Sub[ \t]+Auto_Open\b", Options);

    /// <summary>A procedure only the licensing module should define: a leftover copy would clash or run instead.</summary>
    private static readonly Regex LicensingProcedure = new(
        @"^[ \t]*(?:(?:Public|Private|Friend)[ \t]+)?(?:Static[ \t]+)?(?:Function|Sub|Property[ \t]+(?:Get|Let|Set))[ \t]+(Cfg_ProductKey|License_\w+)\b", Options);

    /// <summary>One element of the saved Cfg_ProtectedSheetNames array, as LicenseModuleConfigurator writes it.</summary>
    private static readonly Regex ProtectedNameLine = new(@"^[ \t]*protectedNames\((\d+)\)[ \t]*=[ \t]*(.*?)[ \t]*$", Options);

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

    /// <summary>
    /// The product's own modules (outside what the package manages) must not define Auto_Open, which is the
    /// licensing module's gate, nor Cfg_ProductKey or License_* (an old copy of the licensing module).
    /// </summary>
    public static string? FindForeignLicensingCode(
        IVbaWorkbookSession session, IReadOnlyList<VbaComponentInfo> components, IReadOnlyCollection<string> replaced)
    {
        var replacedNames = replaced.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var modules = components.Where(c => c.Kind is VbaComponentKind.Standard or VbaComponentKind.Class && !replacedNames.Contains(c.Name));
        foreach (var module in modules)
        {
            var code = session.GetDocumentCode(module.Name);
            var licensing = LicensingProcedure.Match(code);
            if (licensing.Success)
                return $"Module {module.Name} của file không do tool-vba quản lý nhưng có thủ tục {licensing.Groups[1].Value} của module bản quyền " +
                       "(có thể là một bản module bản quyền cũ). Hãy xóa module đó trong file gốc (Alt+F11) rồi thử lại.";
            if (module.Kind == VbaComponentKind.Standard && AutoOpen.IsMatch(code))
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

    /// <summary>
    /// Checks the saved copy against what was packaged: the guide is the first sheet and visible, every other sheet
    /// is very hidden and is exactly the protected list, and the module carries the product, the version and that
    /// same list. Null when all is right.
    /// </summary>
    public static string? VerifySaved(IVbaWorkbookSession session, LicenseSetup setup)
    {
        var code = session.GetDocumentCode(setup.Module.Name);
        if (!code.Contains($"Cfg_ProductKey = \"{setup.Product.ProductKey}\"", StringComparison.Ordinal) ||
            !code.Contains($"Cfg_ProductVersion = \"{setup.Product.Version}\"", StringComparison.Ordinal))
            return "Module bản quyền trong file tạm chưa có mã sản phẩm và phiên bản.";
        var expected = setup.Sheets.ProtectedNames;
        var saved = VbaSourceRules.SplitLines(code).Select(l => ProtectedNameLine.Match(l)).Where(m => m.Success)
            .Select(m => (Index: int.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture), Literal: m.Groups[2].Value))
            .ToList();
        if (!saved.SequenceEqual(expected.Select((name, i) => (i, VbaStringLiteral.From(name)))))
            return "Danh sách sheet bảo vệ trong module bản quyền đã lưu khác danh sách sheet đã ẩn.";

        var sheets = session.GetSheets();
        if (sheets.Count == 0 || !string.Equals(sheets[0].Name, GuideSheet.Name, StringComparison.OrdinalIgnoreCase) ||
            sheets[0].Visibility != SheetVisibility.Visible)
            return "Sheet \"Huong dan\" phải là sheet thứ nhất và đang hiện trong file tạm.";
        var exposed = sheets.Skip(1).FirstOrDefault(s => s.Visibility != SheetVisibility.VeryHidden);
        if (exposed is not null) return $"Sheet \"{exposed.Name}\" chưa được ẩn trong file tạm.";
        return sheets.Skip(1).Select(s => s.Name).SequenceEqual(expected, StringComparer.Ordinal)
            ? null
            : "Các sheet trong file tạm khác danh sách sheet đã ẩn.";
    }
}
