namespace ToolVba.Core;

/// <summary>
/// Removes full-line comments and blank lines from VBA source as a casual-reader deterrent (DEC-116):
/// the injected licensing module keeps working exactly the same, but the narrated "why" an attacker
/// would lean on is gone. This never renames identifiers and never touches a line that carries code,
/// so the exact-string contracts of <see cref="LicenseModuleConfigurator"/> and
/// <see cref="LicensePackaging.VerifySaved"/> still hold. It is deliberately conservative:
/// <list type="bullet">
/// <item>Only whole-line <c>'</c> and <c>Rem</c> comments and whitespace-only lines are dropped; a
/// trailing (inline) comment is kept, so a <c>'</c> inside a string literal is never misread.</item>
/// <item>A physical line that continues the previous one (the previous kept line ends with
/// <c> _</c>) is always kept, so a multi-line statement such as a <c>Declare</c> is never split.</item>
/// </list>
/// </summary>
public static class VbaSourceMinifier
{
    /// <summary>
    /// Returns <paramref name="source"/> with whole-line comments and blank lines removed. Lines are
    /// rejoined with CRLF and the result ends with a trailing CRLF when anything remains.
    /// </summary>
    public static string StripCommentsAndBlankLines(string source)
    {
        ArgumentNullException.ThrowIfNull(source);
        var kept = new List<string>();
        var keepNext = false;
        foreach (var line in VbaSourceRules.SplitLines(source))
        {
            if (keepNext)
            {
                // Inside a line-continuation run: this physical line belongs to the statement above.
                kept.Add(line);
                keepNext = ContinuesOntoNextLine(line);
                continue;
            }

            var trimmed = line.TrimStart();
            if (trimmed.Length == 0 || IsWholeLineComment(trimmed))
                continue;

            kept.Add(line);
            keepNext = ContinuesOntoNextLine(line);
        }
        return kept.Count == 0 ? "" : string.Join("\r\n", kept) + "\r\n";
    }

    private static bool IsWholeLineComment(string trimmedStart) =>
        trimmedStart[0] == '\'' ||
        (trimmedStart.StartsWith("Rem", StringComparison.OrdinalIgnoreCase) &&
         (trimmedStart.Length == 3 || char.IsWhiteSpace(trimmedStart[3])));

    /// <summary>True when the line ends with a VBA line-continuation: whitespace then a trailing underscore.</summary>
    private static bool ContinuesOntoNextLine(string line)
    {
        var trimmedEnd = line.TrimEnd();
        return trimmedEnd.Length >= 2 && trimmedEnd[^1] == '_' && char.IsWhiteSpace(trimmedEnd[^2]);
    }
}
