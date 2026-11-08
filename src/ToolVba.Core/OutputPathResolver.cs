using System.Text.RegularExpressions;

namespace ToolVba.Core;

/// <summary>Chooses <c>&lt;name&gt;_vba.xlsm</c> next to the source, numbering it when the name is taken.</summary>
public static class OutputPathResolver
{
    public const string Suffix = "_vba";
    private const int MaxNumberedAttempts = 999;
    // Strips only a suffix this tool added ("_vba" or "_vba (3)"); "Report (2024)" keeps its own number.
    private static readonly Regex ExistingSuffix = new(@"^(?<stem>.*?)(?:_vba(?: \(\d+\))?)?$", RegexOptions.CultureInvariant);

    /// <param name="inputPath">Source workbook.</param>
    /// <param name="directory">Output folder; <c>null</c> means next to the source.</param>
    /// <param name="reserved">Full paths already chosen for other files in the same batch.</param>
    public static string Resolve(string inputPath, string? directory = null, IReadOnlySet<string>? reserved = null)
    {
        var input = Path.GetFullPath(inputPath);
        var folder = directory is null ? Path.GetDirectoryName(input)! : Path.GetFullPath(directory);
        var stem = ExistingSuffix.Match(Path.GetFileNameWithoutExtension(input)).Groups["stem"].Value;

        for (var number = 1; number <= MaxNumberedAttempts; number++)
        {
            var candidate = Path.Combine(folder, stem + Suffix + (number == 1 ? "" : $" ({number})") + ".xlsm");
            if (!File.Exists(candidate) && reserved?.Contains(candidate) != true &&
                !string.Equals(candidate, input, StringComparison.OrdinalIgnoreCase))
                return candidate;
        }
        return Path.Combine(folder, $"{stem}{Suffix} ({Guid.NewGuid():N}).xlsm");
    }
}
