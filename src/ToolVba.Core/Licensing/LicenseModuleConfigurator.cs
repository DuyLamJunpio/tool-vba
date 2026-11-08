namespace ToolVba.Core;

/// <summary>
/// Fills the per-product settings of a copy of the licensing module (modLicenseApiV1.bas): the published
/// source leaves <c>Cfg_ProductKey</c>, <c>Cfg_ProductVersion</c> and <c>Cfg_ProtectedSheetNames</c> empty
/// and the packager's app writes them at "Tạo" (DEC-112, DEC-115). Exactly one empty line of each setting
/// must exist, so a changed or already-filled module is refused instead of being half-configured.
/// </summary>
public static class LicenseModuleConfigurator
{
    private const string KeyLine = "Cfg_ProductKey = vbNullString";
    private const string VersionLine = "Cfg_ProductVersion = vbNullString";
    private const string SheetsLine = "Cfg_ProtectedSheetNames = Array()";
    private const string ArrayName = "protectedNames";

    private static readonly char[] ForbiddenSheetCharacters = ['\\', '/', '?', '*', '[', ']', ':'];

    public static (string? Code, string? Error) Configure(
        string moduleSource, LicenseProduct product, IReadOnlyList<string> protectedSheetNames)
    {
        ArgumentNullException.ThrowIfNull(moduleSource);
        ArgumentNullException.ThrowIfNull(product);
        ArgumentNullException.ThrowIfNull(protectedSheetNames);
        // The record can be built without its factory; the key and version end up inside VBA code.
        var (checkedProduct, productError) = LicenseProduct.FromValues(product.ProductKey, product.Version);
        if (checkedProduct is null) return (null, productError);
        var sheetError = CheckSheetNames(protectedSheetNames);
        if (sheetError is not null) return (null, sheetError);
        if (!moduleSource.All(char.IsAscii))
            return (null, "File module bản quyền có ký tự không phải ASCII. Hãy xuất bản lại gói từ bản module gốc.");

        var lines = VbaSourceRules.SplitLines(moduleSource).ToList();
        var replacements = new (string Placeholder, Func<string, IEnumerable<string>> Write)[]
        {
            (KeyLine, indent => [indent + "Cfg_ProductKey = " + VbaStringLiteral.From(checkedProduct.ProductKey)]),
            (VersionLine, indent => [indent + "Cfg_ProductVersion = " + VbaStringLiteral.From(checkedProduct.Version)]),
            (SheetsLine, indent => SheetArrayLines(indent, protectedSheetNames))
        };
        foreach (var (placeholder, write) in replacements)
        {
            var matches = lines.Select((line, index) => (line, index)).Where(l => l.line.Trim() == placeholder).ToList();
            if (matches.Count != 1)
                return (null, $"File module bản quyền không đúng mẫu: cần đúng một dòng '{placeholder}'. " +
                              "Hãy xuất bản lại gói từ file mẫu có ba dòng cấu hình để trống.");
            var (line, index) = matches[0];
            lines.RemoveAt(index);
            lines.InsertRange(index, write(line[..^line.TrimStart().Length]));
        }
        return (string.Join("\r\n", lines), null);
    }

    /// <summary>One assignment per line: a long list or a name full of ChrW() never reaches VBA's line limit.</summary>
    private static IEnumerable<string> SheetArrayLines(string indent, IReadOnlyList<string> names)
    {
        yield return $"{indent}Dim {ArrayName}(0 To {names.Count - 1}) As Variant";
        for (var i = 0; i < names.Count; i++)
            yield return $"{indent}{ArrayName}({i}) = {VbaStringLiteral.From(names[i])}";
        yield return $"{indent}Cfg_ProtectedSheetNames = {ArrayName}";
    }

    /// <summary>Excel's own sheet-name rules, so a name the module will look up can exist in the workbook.</summary>
    private static string? CheckSheetNames(IReadOnlyList<string> names)
    {
        if (names.Count == 0)
            return "Không có sheet nào để bảo vệ: file cần ít nhất một sheet ngoài sheet hướng dẫn.";
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in names)
        {
            if (!IsValidSheetName(name)) return $"Tên sheet '{name}' không hợp lệ trong Excel.";
            if (!seen.Add(name)) return $"Tên sheet '{name}' bị trùng.";
        }
        return null;
    }

    private static bool IsValidSheetName(string? name) =>
        !string.IsNullOrWhiteSpace(name) &&
        name.Length <= VbaSourceRules.MaxSheetNameLength &&
        name.IndexOfAny(ForbiddenSheetCharacters) < 0 &&
        !name.Any(char.IsControl) &&
        !name.StartsWith('\'') && !name.EndsWith('\'') &&
        // Excel reserves this name for its own change-history sheet.
        !name.Equals("History", StringComparison.OrdinalIgnoreCase);
}
