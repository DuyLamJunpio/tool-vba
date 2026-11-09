using System.Globalization;
using System.Text;

namespace ToolVba.Core;

/// <summary>
/// Writes text as a VBA string expression made of printable ASCII only, so a module that must stay
/// ASCII can still hold names such as "Dữ liệu": quoted runs for printable ASCII (with <c>"</c> doubled)
/// and <c>ChrW(n)</c> for every other UTF-16 code unit, joined with <c>&amp;</c>.
/// The result is one physical line of up to about 13 characters per input character: callers keep the
/// input short enough for VBA's 1023-character line limit (sheet names are at most 31 characters).
/// </summary>
public static class VbaStringLiteral
{
    public static string From(string value)
    {
        if (value.Length == 0) return "\"\"";
        var parts = new List<string>();
        var run = new StringBuilder();
        foreach (var ch in value)
        {
            if (ch is >= ' ' and <= '~')
            {
                run.Append(ch == '"' ? "\"\"" : ch.ToString());
                continue;
            }
            FlushRun(parts, run);
            parts.Add("ChrW(" + ((int)ch).ToString(CultureInfo.InvariantCulture) + ")");
        }
        FlushRun(parts, run);
        return string.Join(" & ", parts);
    }

    private static void FlushRun(List<string> parts, StringBuilder run)
    {
        if (run.Length == 0) return;
        parts.Add("\"" + run + "\"");
        run.Clear();
    }
}
