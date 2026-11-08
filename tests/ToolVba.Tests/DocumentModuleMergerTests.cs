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
    public void Hook_only_output_still_counts_as_managed_code_and_merges_again_unchanged()
    {
        const string product =
            "Option Explicit\r\nPrivate mUser As Long\r\n\r\nPublic Sub Helper()\r\nEnd Sub\r\n" +
            "Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n    mUser = 1\r\nEnd Sub\r\n" +
            "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\r\n    mUser = 2\r\nEnd Sub\r\n" +
            "Private Sub Workbook_AfterSave(ByVal Success As Boolean)\r\n    mUser = 3\r\nEnd Sub\r\n";

        var first = DocumentModuleMerger.Merge(product, LicenseEvents);
        var second = DocumentModuleMerger.Merge(first.Code!, LicenseEvents);

        Assert.True(first.Succeeded, first.Error);
        // VerifySaved relies on this when every licensing event went into a product handler.
        Assert.True(DocumentModuleMerger.ContainsManagedCode(first.Code!));
        Assert.Equal(first.Code, second.Code);
    }

    [Fact]
    public void Merges_the_shipped_licensing_events_into_existing_handlers()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "ToolVba.sln")))
            directory = directory.Parent;
        var events = File.ReadAllText(Path.Combine(directory!.FullName, "license-sample", "api-v1-one-file", "ThisWorkbook-events.txt"));

        var result = DocumentModuleMerger.Merge(ProductWithSaveHandler, events);

        Assert.True(result.Succeeded, result.Error);
        Assert.Contains("    If Not modLicenseApiV1.License_BeforeSave() Then Cancel = True: Exit Sub", result.Code);
    }

    [Theory]
    [InlineData("    Dim done As Boolean")]
    [InlineData("    On Error Resume Next")]
    [InlineData("    GoTo Finish")]
    [InlineData("Finish:")]
    public void Refuses_a_package_event_body_that_would_change_the_product_handler(string line)
    {
        var package = LicenseEvents.Replace("    modLicenseApiV1.License_OnClose\r\n", $"{line}\r\n    modLicenseApiV1.License_OnClose\r\n");
        const string product = "Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n    DoEvents\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(product, package);

        Assert.False(result.Succeeded);
        Assert.Contains("Workbook_BeforeClose", result.Error);
    }

    // Owner, 08/11/2026: in BeforeClose the licensing call goes LAST and only when the product did not cancel
    // (a product asking "close?" and the user answering No must keep its sheets and access).
    private const string ProductWithCloseHandler =
        "Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n" +
        "    If MsgBox(\"Ban chac chan muon dong?\", vbYesNo) = vbNo Then Cancel = True\r\n" +
        "End Sub\r\n";

    [Fact]
    public void Puts_the_close_call_last_and_only_when_the_product_did_not_cancel()
    {
        var result = DocumentModuleMerger.Merge(ProductWithCloseHandler, LicenseEvents);

        Assert.True(result.Succeeded, result.Error);
        var lines = result.Code!.Split("\r\n").ToList();
        var signature = lines.IndexOf("Private Sub Workbook_BeforeClose(Cancel As Boolean)");
        Assert.Equal("    If MsgBox(\"Ban chac chan muon dong?\", vbYesNo) = vbNo Then Cancel = True", lines[signature + 1]);
        Assert.Equal(DocumentModuleMerger.HookBegin, lines[signature + 2]);
        Assert.Equal("    If Not Cancel Then modLicenseApiV1.License_OnClose", lines[signature + 3]);
        Assert.Equal(DocumentModuleMerger.HookEnd, lines[signature + 4]);
        Assert.Equal("End Sub", lines[signature + 5]);
    }

    [Fact]
    public void Wraps_a_close_body_of_several_lines_in_one_if_block()
    {
        var package = LicenseEvents.Replace("    modLicenseApiV1.License_OnClose\r\n", "    modLicenseApiV1.License_OnClose\r\n    modLicenseApiV1.License_OnClose\r\n");

        var result = DocumentModuleMerger.Merge(ProductWithCloseHandler, package);

        Assert.True(result.Succeeded, result.Error);
        Assert.Contains("    If Not Cancel Then\r\n        modLicenseApiV1.License_OnClose\r\n        modLicenseApiV1.License_OnClose\r\n    End If\r\n", result.Code);
    }

    [Fact]
    public void Does_not_count_a_comment_as_the_one_close_statement()
    {
        // "If Not Cancel Then ' note" would be a block If without End If.
        var package = LicenseEvents.Replace("    modLicenseApiV1.License_OnClose\r\n", "    ' note\r\n    modLicenseApiV1.License_OnClose\r\n");

        var result = DocumentModuleMerger.Merge(ProductWithCloseHandler, package);

        Assert.True(result.Succeeded, result.Error);
        Assert.Contains("    If Not Cancel Then modLicenseApiV1.License_OnClose\r\n", result.Code);
        Assert.DoesNotContain("If Not Cancel Then '", result.Code);
    }

    [Fact]
    public void Packaging_again_keeps_a_single_close_call_at_the_end()
    {
        var first = DocumentModuleMerger.Merge(ProductWithCloseHandler, LicenseEvents).Code!;

        var second = DocumentModuleMerger.Merge(first, LicenseEvents);

        Assert.True(second.Succeeded, second.Error);
        Assert.Equal(first, second.Code);
    }

    [Theory]
    [InlineData("    If Cancel Then Exit Sub")]
    [InlineData("    Exit Sub")]
    [InlineData("    Exit Function")]
    [InlineData("    End")]
    [InlineData("    If ThisWorkbook.Saved Then End")]
    [InlineData("    x = 1: End")]
    public void Refuses_a_close_handler_that_can_leave_before_the_licensing_call(string line)
    {
        var product = $"Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n{line}\r\n    DoEvents\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(product, LicenseEvents);

        Assert.False(result.Succeeded);
        Assert.Contains("Workbook_BeforeClose", result.Error);
        Assert.Contains("Exit Sub", result.Error);
    }

    [Theory]
    [InlineData("    If Cancel Then\r\n        DoEvents\r\n    End If")]
    [InlineData("    With ThisWorkbook\r\n        .Saved = True\r\n    End With")]
    [InlineData("    ' Exit Sub is not used here")]
    [InlineData("    MsgBox \"End\"")]
    public void Accepts_block_endings_comments_and_strings_in_a_close_handler(string lines)
    {
        var product = $"Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n{lines}\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(product, LicenseEvents);

        Assert.True(result.Succeeded, result.Error);
    }

    [Fact]
    public void Exit_sub_in_a_save_handler_is_allowed_because_its_call_runs_first()
    {
        var product = "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\r\n    If SaveAsUI Then Exit Sub\r\nEnd Sub\r\n";

        var result = DocumentModuleMerger.Merge(product, LicenseEvents);

        Assert.True(result.Succeeded, result.Error);
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
