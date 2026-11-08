using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class DocumentModuleMergerTests
{
    private const string PackageCode =
        "Option Explicit\r\n" +
        "Private mLoaded As Boolean\r\n" +
        "\r\n" +
        "Private Sub Workbook_Open()\r\n" +
        "    mLoaded = True\r\n" +
        "End Sub\r\n";

    [Fact]
    public void Adds_declarations_after_user_declarations_and_code_at_the_end()
    {
        const string existing =
            "Option Explicit\r\n" +
            "Private mUser As Long\r\n" +
            "\r\n" +
            "Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n" +
            "End Sub\r\n";

        var result = DocumentModuleMerger.Merge(existing, PackageCode);

        Assert.True(result.Succeeded, result.Error);
        var lines = result.Code!.Split("\r\n");
        var userDecl = Array.IndexOf(lines, "Private mUser As Long");
        var declBegin = Array.IndexOf(lines, DocumentModuleMerger.DeclarationsBegin);
        var managedDecl = Array.IndexOf(lines, "Private mLoaded As Boolean");
        var userProc = Array.IndexOf(lines, "Private Sub Workbook_BeforeClose(Cancel As Boolean)");
        var codeBegin = Array.IndexOf(lines, DocumentModuleMerger.CodeBegin);
        var managedProc = Array.IndexOf(lines, "Private Sub Workbook_Open()");
        Assert.True(userDecl < declBegin && declBegin < managedDecl && managedDecl < userProc);
        Assert.True(userProc < codeBegin && codeBegin < managedProc);
        Assert.Single(lines, line => line == "Option Explicit");
    }

    [Fact]
    public void Replaces_previous_managed_regions_instead_of_duplicating_them()
    {
        var first = DocumentModuleMerger.Merge("Option Explicit\r\n", PackageCode).Code!;
        var updatedPackage = PackageCode.Replace("mLoaded = True", "mLoaded = False");

        var second = DocumentModuleMerger.Merge(first, updatedPackage);

        Assert.True(second.Succeeded, second.Error);
        Assert.Single(second.Code!.Split("\r\n"), line => line == DocumentModuleMerger.CodeBegin);
        Assert.Contains("mLoaded = False", second.Code);
        Assert.DoesNotContain("mLoaded = True", second.Code);
    }

    // DEC-115: the licensing events go at the top of the product's own handlers instead of refusing.
    private const string LicenseEvents =
        "Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n" +
        "    modLicenseApiV1.License_OnClose\r\n" +
        "End Sub\r\n" +
        "\r\n" +
        "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\r\n" +
        "    If Not modLicenseApiV1.License_BeforeSave() Then Cancel = True: Exit Sub\r\n" +
        "End Sub\r\n" +
        "\r\n" +
        "Private Sub Workbook_AfterSave(ByVal Success As Boolean)\r\n" +
        "    modLicenseApiV1.License_AfterSave\r\n" +
        "End Sub\r\n";

    private const string ProductWithSaveHandler =
        "Option Explicit\r\n" +
        "\r\n" +
        "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\r\n" +
        "    Sheet1.Range(\"A1\").Value = Now\r\n" +
        "End Sub\r\n";

    [Fact]
    public void Puts_licensing_calls_at_the_top_of_an_existing_event_handler()
    {
        var result = DocumentModuleMerger.Merge(ProductWithSaveHandler, LicenseEvents);

        Assert.True(result.Succeeded, result.Error);
        var lines = result.Code!.Split("\r\n").ToList();
        var signature = lines.IndexOf("Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)");
        Assert.Equal(DocumentModuleMerger.HookBegin, lines[signature + 1]);
        Assert.Equal("    If Not modLicenseApiV1.License_BeforeSave() Then Cancel = True: Exit Sub", lines[signature + 2]);
        Assert.Equal(DocumentModuleMerger.HookEnd, lines[signature + 3]);
        Assert.Equal("    Sheet1.Range(\"A1\").Value = Now", lines[signature + 4]);
        // The product's handler is the only one; the other two stay in the managed region.
        Assert.Single(lines, l => l.Contains("Sub Workbook_BeforeSave"));
        var codeBegin = lines.IndexOf(DocumentModuleMerger.CodeBegin);
        Assert.True(codeBegin > signature);
        Assert.True(lines.IndexOf("Private Sub Workbook_BeforeClose(Cancel As Boolean)") > codeBegin);
        Assert.True(lines.IndexOf("Private Sub Workbook_AfterSave(ByVal Success As Boolean)") > codeBegin);
    }

    [Fact]
    public void Packaging_again_replaces_the_calls_instead_of_adding_them_twice()
    {
        var first = DocumentModuleMerger.Merge(ProductWithSaveHandler, LicenseEvents).Code!;

        var second = DocumentModuleMerger.Merge(first, LicenseEvents);

        Assert.True(second.Succeeded, second.Error);
        Assert.Equal(first, second.Code);
    }

    [Fact]
    public void Merges_into_all_three_handlers_and_keeps_each_handler_body()
    {
        const string product =
            "Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n    SaveSetting \"a\", \"b\", \"c\", \"d\"\r\nEnd Sub\r\n" +
            "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\r\n    DoEvents\r\nEnd Sub\r\n" +
            "Private Sub Workbook_AfterSave(ByVal Success As Boolean)\r\n    Beep\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(product, LicenseEvents);

        Assert.True(result.Succeeded, result.Error);
        Assert.Equal(3, result.Code!.Split("\r\n").Count(l => l == DocumentModuleMerger.HookBegin));
        Assert.DoesNotContain(DocumentModuleMerger.CodeBegin, result.Code);
        Assert.Contains("    SaveSetting \"a\", \"b\", \"c\", \"d\"", result.Code);
        Assert.Contains("    DoEvents", result.Code);
        Assert.Contains("    Beep", result.Code);
    }

    [Theory]
    [InlineData("Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Huy As Boolean)")]
    [InlineData("Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, _\r\n    Cancel As Boolean)")]
    [InlineData("Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean) ' mine")]
    public void Stops_with_a_clear_message_when_the_handler_signature_is_unusual(string signature)
    {
        var product = $"{signature}\r\n    DoEvents\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(product, LicenseEvents);

        Assert.False(result.Succeeded);
        Assert.Contains("Workbook_BeforeSave", result.Error);
        Assert.Contains("chữ ký", result.Error);
    }

    [Fact]
    public void Accepts_the_same_signature_written_with_other_spacing_case_or_scope()
    {
        const string product =
            "Sub workbook_beforesave(ByVal   SaveAsUI As Boolean,Cancel As Boolean)\r\n    DoEvents\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(product, LicenseEvents);

        Assert.True(result.Succeeded, result.Error);
        Assert.Contains(DocumentModuleMerger.HookBegin, result.Code);
    }

    [Fact]
    public void Refuses_a_hook_without_its_end_marker()
    {
        var broken = DocumentModuleMerger.Merge(ProductWithSaveHandler, LicenseEvents).Code!
            .Replace(DocumentModuleMerger.HookEnd + "\r\n", "");

        var result = DocumentModuleMerger.Merge(broken, LicenseEvents);

        Assert.False(result.Succeeded);
    }

    [Fact]
    public void Still_refuses_other_procedures_that_exist_in_both()
    {
        const string product = "Private Sub Workbook_AfterSave(ByVal Success As Boolean)\r\nEnd Sub\r\nPublic Sub Helper()\r\nEnd Sub\r\n";
        var package = LicenseEvents + "Public Sub Helper()\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(product, package);

        Assert.False(result.Succeeded);
        Assert.Equal(["Helper"], result.ConflictingNames);
    }

    [Fact]
    public void Reports_procedure_that_already_exists_outside_managed_region()
    {
        const string existing = "Private Sub Workbook_Open()\r\n    MsgBox \"user\"\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(existing, PackageCode);

        Assert.False(result.Succeeded);
        Assert.Contains("Workbook_Open", result.ConflictingNames);
    }

    [Fact]
    public void Reports_module_level_variable_that_already_exists()
    {
        const string existing = "Private mLoaded As Boolean\r\n";

        var result = DocumentModuleMerger.Merge(existing, PackageCode);

        Assert.False(result.Succeeded);
        Assert.Contains("mLoaded", result.ConflictingNames);
    }

    [Fact]
    public void Ignores_declare_statements_when_finding_procedures()
    {
        const string package =
            "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long\r\n" +
            "Public Sub Tick()\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge("", package);

        Assert.True(result.Succeeded, result.Error);
        var lines = result.Code!.Split("\r\n");
        Assert.True(Array.IndexOf(lines, package.Split("\r\n")[0]) < Array.IndexOf(lines, DocumentModuleMerger.CodeBegin));
    }

    [Fact]
    public void Refuses_option_statement_that_would_change_existing_user_code()
    {
        const string package = "Option Base 1\r\nPublic Sub Tick()\r\nEnd Sub\r\n";

        Assert.False(DocumentModuleMerger.Merge("Public Sub Mine()\r\nEnd Sub\r\n", package).Succeeded);
        Assert.StartsWith("Option Base 1", DocumentModuleMerger.Merge("", package).Code);
    }

    [Fact]
    public void Refuses_unterminated_managed_region_instead_of_dropping_code_after_it()
    {
        var existing = DocumentModuleMerger.CodeBegin + "\r\nPublic Sub Old()\r\nEnd Sub\r\nPublic Sub EmployeeMacro()\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(existing, PackageCode);

        Assert.False(result.Succeeded);
        Assert.Contains("END", result.Error);
    }

    [Fact]
    public void Treats_different_values_of_the_same_option_as_a_conflict()
    {
        const string package = "Option Compare Text\r\nPublic Sub Tick()\r\nEnd Sub\r\n";

        var withCode = DocumentModuleMerger.Merge("Option Compare Binary\r\nPublic Sub Mine()\r\nEnd Sub\r\n", package);
        var optionOnly = DocumentModuleMerger.Merge("Option Compare Binary\r\n", package);

        Assert.False(withCode.Succeeded);
        Assert.False(optionOnly.Succeeded);
        Assert.True(DocumentModuleMerger.Merge("Option Compare Text\r\nPublic Sub Mine()\r\nEnd Sub\r\n", package).Succeeded);
    }

    [Fact]
    public void Ignores_marker_lines_inside_package_code_taken_from_a_tool_made_template()
    {
        var templateCode = DocumentModuleMerger.Merge("", PackageCode).Code!;

        var result = DocumentModuleMerger.Merge("", templateCode);

        Assert.True(result.Succeeded, result.Error);
        Assert.Single(result.Code!.Split("\r\n"), line => line == DocumentModuleMerger.CodeBegin);
        Assert.Single(result.Code!.Split("\r\n"), line => line == DocumentModuleMerger.DeclarationsBegin);
    }

    [Fact]
    public void Detects_managed_code_and_empty_modules()
    {
        var merged = DocumentModuleMerger.Merge("", PackageCode).Code!;

        Assert.True(DocumentModuleMerger.ContainsManagedCode(merged));
        Assert.False(DocumentModuleMerger.ContainsManagedCode(PackageCode));
        Assert.True(DocumentModuleMerger.IsEffectivelyEmpty("Option Explicit\r\n\r\n"));
        Assert.False(DocumentModuleMerger.IsEffectivelyEmpty(PackageCode));
    }
}
