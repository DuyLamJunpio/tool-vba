using System.Text.RegularExpressions;

namespace ToolVba.Core;

/// <summary>Static checks shared by every place that accepts VBA source text.</summary>
public static class VbaSourceRules
{
    public const int MaxComponentNameLength = 31;
    public const int MaxSheetNameLength = 31;

    // Excel in Vietnamese names document modules with letters such as "Trang_tính1".
    private static readonly Regex ComponentName = new(
        @"^\p{L}[\p{L}\p{Nd}_]{0,30}$", RegexOptions.CultureInvariant);

    private static readonly Regex AsciiFileStem = new(@"^[A-Za-z0-9_-]{1,40}$", RegexOptions.CultureInvariant);

    private static readonly Regex ReservedInputDeclaration = new(
        @"\b(?:Dim|Const|ByVal|ByRef)\s+Input\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex VbNameAttribute = new(
        @"^\s*Attribute\s+VB_Name\s*=\s*""([^""]+)""", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static bool IsValidComponentName(string? name) => name is not null && ComponentName.IsMatch(name);

    /// <summary>True when the name can be used as-is for a zip entry / file name.</summary>
    public static bool IsAsciiFileStem(string name) => AsciiFileStem.IsMatch(name);

    /// <summary>
    /// Import can succeed even when VBA later rejects a declaration at compile time.
    /// This targeted check prevents the reserved identifier seen in Win32 Declare statements.
    /// </summary>
    public static string? FindReservedInputDeclaration(IEnumerable<string> lines)
    {
        foreach (var line in lines)
        {
            if (!line.TrimStart().StartsWith('\'') && ReservedInputDeclaration.IsMatch(line))
                return "Code dùng từ khóa VBA 'Input' làm tên tham số hoặc biến. Hãy đổi tên rồi xuất bản lại.";
        }
        return null;
    }

    public static string? ReadVbName(IEnumerable<string> lines)
    {
        foreach (var line in lines)
        {
            var match = VbNameAttribute.Match(line);
            if (match.Success) return match.Groups[1].Value;
        }
        return null;
    }

    /// <summary>The only file extension allowed for each component kind.</summary>
    public static string? ExtensionFor(VbaComponentKind kind) => kind switch
    {
        VbaComponentKind.Standard => ".bas",
        VbaComponentKind.Class => ".cls",
        VbaComponentKind.Form => ".frm",
        VbaComponentKind.Workbook or VbaComponentKind.Sheet => ".txt",
        _ => null
    };

    public static string[] SplitLines(string text) => text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
}
