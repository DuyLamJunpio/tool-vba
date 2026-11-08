using System.Text;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class LicenseModuleConfiguratorTests
{
    private static readonly LicenseProduct Product = new("test-01", "1.0.0");

    private const string Template =
        "Attribute VB_Name = \"modLicenseApiV1\"\r\n" +
        "Option Explicit\r\n" +
        "\r\n" +
        "Public Function Cfg_ProductKey() As String\r\n" +
        "    ' Must equal products.product_key for the target workbook.\r\n" +
        "    Cfg_ProductKey = vbNullString\r\n" +
        "End Function\r\n" +
        "\r\n" +
        "Public Function Cfg_ProductVersion() As String\r\n" +
        "    Cfg_ProductVersion = vbNullString\r\n" +
        "End Function\r\n" +
        "\r\n" +
        "Public Function Cfg_ProtectedSheetNames() As Variant\r\n" +
        "    Cfg_ProtectedSheetNames = Array()\r\n" +
        "End Function\r\n";

    [Fact]
    public void Fills_key_version_and_sheet_names_and_keeps_every_other_line()
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, Product, ["Du lieu", "Bao cao"]);

        Assert.Null(error);
        Assert.Equal(
            "Attribute VB_Name = \"modLicenseApiV1\"\r\n" +
            "Option Explicit\r\n" +
            "\r\n" +
            "Public Function Cfg_ProductKey() As String\r\n" +
            "    ' Must equal products.product_key for the target workbook.\r\n" +
            "    Cfg_ProductKey = \"test-01\"\r\n" +
            "End Function\r\n" +
            "\r\n" +
            "Public Function Cfg_ProductVersion() As String\r\n" +
            "    Cfg_ProductVersion = \"1.0.0\"\r\n" +
            "End Function\r\n" +
            "\r\n" +
            "Public Function Cfg_ProtectedSheetNames() As Variant\r\n" +
            "    Dim protectedNames(0 To 1) As Variant\r\n" +
            "    protectedNames(0) = \"Du lieu\"\r\n" +
            "    protectedNames(1) = \"Bao cao\"\r\n" +
            "    Cfg_ProtectedSheetNames = protectedNames\r\n" +
            "End Function\r\n",
            code);
    }

    [Fact]
    public void Writes_accented_sheet_names_with_ChrW_so_the_module_stays_ascii()
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, Product, ["Dữ liệu", "Đầu kỳ"]);

        Assert.Null(error);
        Assert.Contains("    protectedNames(0) = \"D\" & ChrW(7919) & \" li\" & ChrW(7879) & \"u\"\r\n", code, StringComparison.Ordinal);
        Assert.Contains("    protectedNames(1) = ChrW(272) & ChrW(7847) & \"u k\" & ChrW(7923)\r\n", code, StringComparison.Ordinal);
        Assert.True(code!.All(char.IsAscii));
    }

    [Fact]
    public void Configures_the_shipped_license_module()
    {
        var source = File.ReadAllText(ShippedModulePath(), Encoding.Latin1);

        var (code, error) = LicenseModuleConfigurator.Configure(source, Product, ["Du lieu"]);

        Assert.Null(error);
        Assert.Contains("Cfg_ProductKey = \"test-01\"", code, StringComparison.Ordinal);
        Assert.Contains("Cfg_ProductVersion = \"1.0.0\"", code, StringComparison.Ordinal);
        Assert.Contains("protectedNames(0) = \"Du lieu\"", code, StringComparison.Ordinal);
        Assert.DoesNotContain("Cfg_ProductKey = vbNullString", code, StringComparison.Ordinal);
        Assert.True(code!.All(char.IsAscii));
        // One sheet: the single Array() line becomes three lines; nothing else changes.
        Assert.Equal(VbaSourceRules.SplitLines(source).Length + 2, VbaSourceRules.SplitLines(code).Length);
    }

    [Theory]
    [InlineData("    Cfg_ProductKey = vbNullString\r\n")]
    [InlineData("    Cfg_ProductVersion = vbNullString\r\n")]
    [InlineData("    Cfg_ProtectedSheetNames = Array()\r\n")]
    public void Refuses_a_module_without_one_of_the_empty_settings(string line)
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template.Replace(line, ""), Product, ["Du lieu"]);

        Assert.Null(code);
        Assert.Contains("module bản quyền", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("    Cfg_ProductKey = vbNullString")]
    [InlineData("    Cfg_ProductVersion = vbNullString")]
    [InlineData("    Cfg_ProtectedSheetNames = Array()")]
    public void Refuses_a_module_whose_setting_appears_twice(string line)
    {
        var doubled = Template + "Private Sub Again()\r\n" + line + "\r\nEnd Sub\r\n";

        var (code, error) = LicenseModuleConfigurator.Configure(doubled, Product, ["Du lieu"]);

        Assert.Null(code);
        Assert.Contains("module bản quyền", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Refuses_a_setting_line_that_carries_anything_else()
    {
        var commented = Template.Replace("Cfg_ProductKey = vbNullString", "Cfg_ProductKey = vbNullString ' set later");

        var (code, error) = LicenseModuleConfigurator.Configure(commented, Product, ["Du lieu"]);

        Assert.Null(code);
        Assert.Contains("module bản quyền", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Refuses_a_module_already_filled_for_a_product()
    {
        var filled = Template.Replace("Cfg_ProductKey = vbNullString", "Cfg_ProductKey = \"other\"");

        var (code, error) = LicenseModuleConfigurator.Configure(filled, Product, ["Du lieu"]);

        Assert.Null(code);
        Assert.Contains("Cfg_ProductKey = vbNullString", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Writes_crlf_lines_for_a_module_with_unix_or_mixed_line_endings()
    {
        var mixed = Template.Replace("\r\n", "\n").Replace("Option Explicit\n", "Option Explicit\r\n");

        var (code, error) = LicenseModuleConfigurator.Configure(mixed, Product, ["Du lieu"]);

        Assert.Null(error);
        Assert.DoesNotContain('\n', code!.Replace("\r\n", ""));
        Assert.EndsWith("End Function\r\n", code, StringComparison.Ordinal);
    }

    [Fact]
    public void Writes_quotes_and_characters_outside_the_basic_plane_through_configure()
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, Product, ["Bao \"cao\"", "Lịch 😀"]);

        Assert.Null(error);
        Assert.Contains("    protectedNames(0) = \"Bao \"\"cao\"\"\"\r\n", code, StringComparison.Ordinal);
        Assert.Contains("    protectedNames(1) = \"L\" & ChrW(7883) & \"ch \" & ChrW(55357) & ChrW(56832)\r\n", code, StringComparison.Ordinal);
    }

    [Fact]
    public void Accepts_a_sheet_name_of_exactly_31_characters()
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, Product, [new string('x', 31)]);

        Assert.Null(error);
        Assert.NotNull(code);
    }

    [Fact]
    public void Refuses_null_arguments()
    {
        Assert.Throws<ArgumentNullException>(() => LicenseModuleConfigurator.Configure(null!, Product, ["Du lieu"]));
        Assert.Throws<ArgumentNullException>(() => LicenseModuleConfigurator.Configure(Template, null!, ["Du lieu"]));
        Assert.Throws<ArgumentNullException>(() => LicenseModuleConfigurator.Configure(Template, Product, null!));
    }

    [Fact]
    public void Refuses_a_missing_sheet_name_in_the_list()
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, Product, ["Du lieu", null!]);

        Assert.Null(code);
        Assert.Contains("sheet", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Refuses_a_module_that_is_not_plain_ascii()
    {
        var accented = Template.Replace("' Must equal", "' Phải bằng");

        var (code, error) = LicenseModuleConfigurator.Configure(accented, Product, ["Du lieu"]);

        Assert.Null(code);
        Assert.Contains("ASCII", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Refuses_a_product_that_skipped_validation()
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, new LicenseProduct("x\" & Shell(\"calc\") & \"", "1.0.0"), ["Du lieu"]);

        Assert.Null(code);
        Assert.Contains("Mã sản phẩm", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Refuses_a_version_that_skipped_validation()
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, new LicenseProduct("test-01", "1\" & Shell(\"calc\") & \""), ["Du lieu"]);

        Assert.Null(code);
        Assert.Contains("Phiên bản", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Refuses_an_empty_sheet_list()
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, Product, []);

        Assert.Null(code);
        Assert.Contains("sheet", error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("")]
    [InlineData("a/b")]
    [InlineData("a\\b")]
    [InlineData("a?b")]
    [InlineData("a*b")]
    [InlineData("a[b")]
    [InlineData("a]b")]
    [InlineData("a:b")]
    [InlineData("'ab")]
    [InlineData("ab'")]
    [InlineData("a\tb")]
    [InlineData("12345678901234567890123456789012")]
    [InlineData("   ")]
    [InlineData("History")]
    [InlineData("history")]
    public void Refuses_a_name_excel_does_not_allow_for_a_sheet(string name)
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, Product, [name]);

        Assert.Null(code);
        Assert.Contains("sheet", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Refuses_the_same_sheet_twice_in_any_case()
    {
        var (code, error) = LicenseModuleConfigurator.Configure(Template, Product, ["Du lieu", "DU LIEU"]);

        Assert.Null(code);
        Assert.Contains("sheet", error, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("Du lieu", "\"Du lieu\"")]
    [InlineData("Bao \"cao\"", "\"Bao \"\"cao\"\"\"")]
    [InlineData("Ý", "ChrW(221)")]
    [InlineData("Năm 2026", "\"N\" & ChrW(259) & \"m 2026\"")]
    [InlineData("😀", "ChrW(55357) & ChrW(56832)")]
    public void Builds_a_vba_string_expression_from_any_text(string value, string expected)
    {
        Assert.Equal(expected, VbaStringLiteral.From(value));
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
