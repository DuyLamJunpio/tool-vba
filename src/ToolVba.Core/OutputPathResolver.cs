namespace ToolVba.Core;

/// <summary>
/// Chooses <c>&lt;productKey&gt;_&lt;version&gt;_ban-quyen.xlsm</c> (DEC-115) next to the source, or in a chosen folder,
/// numbering it <c>(2)</c>, <c>(3)</c>… when the name is taken. It is never the source file itself.
/// </summary>
public static class OutputPathResolver
{
    public const string Suffix = "_ban-quyen";
    private const int MaxNumberedAttempts = 999;

    /// <param name="inputPath">Source workbook.</param>
    /// <param name="product">The product and version the file is packaged for.</param>
    /// <param name="directory">Output folder; <c>null</c> means next to the source.</param>
    public static string Resolve(string inputPath, LicenseProduct product, string? directory = null)
    {
        var input = Path.GetFullPath(inputPath);
        var folder = directory is null ? Path.GetDirectoryName(input)! : Path.GetFullPath(directory);
        var stem = $"{product.ProductKey}_{product.Version}{Suffix}";

        for (var number = 1; number <= MaxNumberedAttempts; number++)
        {
            var candidate = Path.Combine(folder, stem + (number == 1 ? "" : $" ({number})") + ".xlsm");
            if (!File.Exists(candidate) && !string.Equals(candidate, input, StringComparison.OrdinalIgnoreCase))
                return candidate;
        }
        return Path.Combine(folder, $"{stem} ({Guid.NewGuid():N}).xlsm");
    }
}
