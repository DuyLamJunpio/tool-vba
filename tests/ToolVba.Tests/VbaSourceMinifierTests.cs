using System.Text;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class VbaSourceMinifierTests
{
    [Fact]
    public void Removes_whole_line_comments_and_blank_lines_but_keeps_code()
    {
        var source =
            "Attribute VB_Name = \"m\"\r\n" +
            "' a narrated reason\r\n" +
            "\r\n" +
            "Public Sub Go()\r\n" +
            "    Dim x As Long\r\n" +
            "\r\n" +
            "End Sub\r\n";

        var result = VbaSourceMinifier.StripCommentsAndBlankLines(source);

        Assert.Equal(
            "Attribute VB_Name = \"m\"\r\n" +
            "Public Sub Go()\r\n" +
            "    Dim x As Long\r\n" +
            "End Sub\r\n",
            result);
    }

    [Fact]
    public void Keeps_a_trailing_inline_comment_so_a_quote_inside_a_string_is_never_misread()
    {
        var source =
            "    s = \"it's a value\" ' trailing note\r\n" +
            "    t = 1\r\n";

        var result = VbaSourceMinifier.StripCommentsAndBlankLines(source);

        // The line carries code, so it stays verbatim, apostrophe in the literal and all.
        Assert.Contains("    s = \"it's a value\" ' trailing note\r\n", result, StringComparison.Ordinal);
        Assert.Contains("    t = 1\r\n", result, StringComparison.Ordinal);
    }

    [Fact]
    public void Removes_a_whole_line_rem_comment_but_keeps_an_identifier_that_merely_starts_with_rem()
    {
        var source =
            "Rem a boxed comment\r\n" +
            "    Remarks = 5\r\n" +
            "    REM another\r\n";

        var result = VbaSourceMinifier.StripCommentsAndBlankLines(source);

        Assert.Equal("    Remarks = 5\r\n", result);
    }

    [Fact]
    public void Keeps_every_physical_line_of_a_line_continuation_statement()
    {
        var source =
            "Private Declare PtrSafe Function F Lib \"k\" ( _\r\n" +
            "    ByVal a As Long, _\r\n" +
            "    ByVal b As Long) As Long\r\n";

        var result = VbaSourceMinifier.StripCommentsAndBlankLines(source);

        Assert.Equal(source, result);
    }

    [Fact]
    public void Returns_empty_for_a_source_that_is_all_comments_and_blanks()
    {
        var result = VbaSourceMinifier.StripCommentsAndBlankLines("' only\r\n\r\nRem nothing\r\n");

        Assert.Equal("", result);
    }

    [Fact]
    public void Handles_unix_and_mixed_line_endings_and_normalizes_to_crlf()
    {
        var result = VbaSourceMinifier.StripCommentsAndBlankLines("Public Sub A()\n' c\nEnd Sub\n");

        Assert.Equal("Public Sub A()\r\nEnd Sub\r\n", result);
    }

    [Fact]
    public void Throws_for_a_null_source()
    {
        Assert.Throws<ArgumentNullException>(() => VbaSourceMinifier.StripCommentsAndBlankLines(null!));
    }

    [Fact]
    public void Stripping_the_configured_shipped_module_keeps_the_verify_contract_and_stays_ascii()
    {
        // Mirror LicensePackaging.Prepare: configure first, then strip the result.
        var source = File.ReadAllText(ShippedModulePath(), Encoding.Latin1);
        var (configured, error) = LicenseModuleConfigurator.Configure(source, new LicenseProduct("test-01", "1.0.0"), ["Du lieu"]);
        Assert.Null(error);

        var stripped = VbaSourceMinifier.StripCommentsAndBlankLines(configured!);

        // The three strings LicensePackaging.VerifySaved reads back must survive stripping.
        Assert.Contains("Cfg_ProductKey = \"test-01\"", stripped, StringComparison.Ordinal);
        Assert.Contains("Cfg_ProductVersion = \"1.0.0\"", stripped, StringComparison.Ordinal);
        Assert.Contains("protectedNames(0) = \"Du lieu\"", stripped, StringComparison.Ordinal);
        // The module keeps its name line and import stays valid; it is still ASCII-only.
        Assert.StartsWith("Attribute VB_Name = \"modLicenseApiV1\"\r\n", stripped, StringComparison.Ordinal);
        Assert.True(stripped.All(char.IsAscii));
        // No whole-line comment remains.
        Assert.DoesNotContain(VbaSourceRules.SplitLines(stripped), l => l.TrimStart().StartsWith('\''));
    }

    private static string ShippedModulePath()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "ToolVba.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return Path.Combine(directory!.FullName, "license-sample", "api-v1-one-file", "modLicenseApiV1.bas");
    }
}
