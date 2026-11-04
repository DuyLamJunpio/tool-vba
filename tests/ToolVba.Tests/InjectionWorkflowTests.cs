using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class InjectionWorkflowTests
{
    [Fact]
    public async Task Injects_into_xlsx_and_produces_macro_enabled_package_with_matching_vba_hash()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("source.xlsx");
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));
        var output = temp.PathFor("result.xlsm");
        WorkbookFixture.CreateWorkbook(source, SpreadsheetDocumentType.Workbook);
        var sourceHash = WorkbookFixture.HashFile(source);

        var result = await new FileOperationService().CreateAsync(Request(source, vba, output));

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(source));
        AssertPackage(output, vba);
    }

    [Fact]
    public async Task Injects_into_xlsm_that_does_not_yet_contain_vba()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("source.xlsm");
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));
        var output = temp.PathFor("result.xlsm");
        WorkbookFixture.CreateWorkbook(source, SpreadsheetDocumentType.MacroEnabledWorkbook);

        var result = await new FileOperationService().CreateAsync(Request(source, vba, output));

        Assert.True(result.Succeeded, result.Message);
        AssertPackage(output, vba);
    }

    [Fact]
    public async Task Refuses_existing_vba_without_explicit_replacement_confirmation()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("source.xlsm");
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));
        var output = temp.PathFor("result.xlsm");
        WorkbookFixture.CreateWorkbook(source, SpreadsheetDocumentType.MacroEnabledWorkbook, withVba: true);
        var originalHash = WorkbookFixture.HashFile(source);

        var result = await new FileOperationService().CreateAsync(Request(source, vba, output));

        Assert.False(result.Succeeded);
        Assert.Contains("Thay VBA", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(originalHash, WorkbookFixture.HashFile(source));
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task Replaces_existing_vba_only_after_explicit_confirmation()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("source.xlsm");
        var newVba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"), [9, 8, 7, 6]);
        var output = temp.PathFor("result.xlsm");
        WorkbookFixture.CreateWorkbook(source, SpreadsheetDocumentType.MacroEnabledWorkbook, withVba: true);
        var sourceHash = WorkbookFixture.HashFile(source);

        var result = await new FileOperationService().CreateAsync(Request(source, newVba, output) with { ConfirmVbaReplacement = true });

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(source));
        AssertPackage(output, newVba);
    }

    [Theory]
    [InlineData("legacy.xls")]
    [InlineData("binary.xlsb")]
    [InlineData("word.docx")]
    [InlineData("other.pptx")]
    public async Task Rejects_unsupported_file_extensions(string name)
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor(name);
        File.WriteAllBytes(source, [1, 2, 3]);
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));

        var result = await new FileOperationService().CreateAsync(Request(source, vba, temp.PathFor("result.xlsm")));

        Assert.False(result.Succeeded);
        Assert.Contains(".xlsx", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Rejects_non_zip_or_invalid_open_xml_input()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("not-a-workbook.xlsx");
        File.WriteAllBytes(source, [1, 2, 3, 4]);
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));

        var result = await new FileOperationService().CreateAsync(Request(source, vba, temp.PathFor("result.xlsm")));

        Assert.False(result.Succeeded);
        Assert.False(Directory.EnumerateFiles(temp.Root, ".tool-vba-*.tmp").Any());
    }

    [Fact]
    public async Task Rejects_xlsx_that_already_contains_vba_instead_of_treating_it_as_macro_enabled()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("mismatched.xlsx");
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));
        WorkbookFixture.CreateWorkbook(source, SpreadsheetDocumentType.Workbook, withVba: true);

        var result = await new FileOperationService().CreateAsync(Request(source, vba, temp.PathFor("result.xlsm")));

        Assert.False(result.Succeeded);
        Assert.Contains(".xlsx", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Does_not_overwrite_existing_output_without_confirmation()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("source.xlsx");
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));
        var output = temp.PathFor("result.xlsm");
        WorkbookFixture.CreateWorkbook(source, SpreadsheetDocumentType.Workbook);
        File.WriteAllBytes(output, [42, 43, 44]);

        var result = await new FileOperationService().CreateAsync(Request(source, vba, output));

        Assert.False(result.Succeeded);
        Assert.Equal(new byte[] { 42, 43, 44 }, await File.ReadAllBytesAsync(output));
    }

    [Fact]
    public async Task Preserves_existing_output_until_replacement_is_confirmed_and_succeeds()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("source.xlsx");
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));
        var output = temp.PathFor("result.xlsm");
        WorkbookFixture.CreateWorkbook(source, SpreadsheetDocumentType.Workbook);
        File.WriteAllBytes(output, [42, 43, 44]);

        var result = await new FileOperationService().CreateAsync(Request(source, vba, output) with { ConfirmOutputOverwrite = true });

        Assert.True(result.Succeeded, result.Message);
        AssertPackage(output, vba);
    }

    [Fact]
    public async Task Keeps_input_and_existing_output_and_removes_temporary_file_after_failure()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("broken.xlsx");
        File.WriteAllBytes(source, [11, 22, 33]);
        var sourceHash = WorkbookFixture.HashFile(source);
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));
        var output = temp.PathFor("result.xlsm");
        var priorOutput = new byte[] { 42, 43, 44 };
        await File.WriteAllBytesAsync(output, priorOutput);

        var result = await new FileOperationService().CreateAsync(Request(source, vba, output) with { ConfirmOutputOverwrite = true });

        Assert.False(result.Succeeded);
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(source));
        Assert.Equal(priorOutput, await File.ReadAllBytesAsync(output));
        Assert.False(Directory.EnumerateFiles(temp.Root, ".tool-vba-*.tmp").Any());
    }

    [Fact]
    public async Task Rejects_output_path_that_is_the_same_as_input()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("same.xlsx");
        var vba = WorkbookFixture.WriteVba(temp.PathFor("approved.bin"));
        WorkbookFixture.CreateWorkbook(source, SpreadsheetDocumentType.Workbook);

        var result = await new FileOperationService().CreateAsync(Request(source, vba, source));

        Assert.False(result.Succeeded);
        Assert.Contains("trùng", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Supports_unicode_names_and_paths_containing_spaces()
    {
        using var temp = new TestDirectory("kiểm thử đường dẫn có dấu");
        var source = temp.PathFor("bảng tính mẫu.xlsx");
        var vba = WorkbookFixture.WriteVba(temp.PathFor("dự án VBA.bin"));
        var output = temp.PathFor("kết quả macro.xlsm");
        WorkbookFixture.CreateWorkbook(source, SpreadsheetDocumentType.Workbook);

        var result = await new FileOperationService().CreateAsync(Request(source, vba, output));

        Assert.True(result.Succeeded, result.Message);
        AssertPackage(output, vba);
    }

    private static InjectionRequest Request(string input, string vba, string output) => new(input, vba, output);

    private static void AssertPackage(string output, string vba)
    {
        var validation = new ExcelPackageValidator().Validate(output, vba);
        Assert.True(validation.IsValid, string.Join("; ", validation.Errors));
        using var document = SpreadsheetDocument.Open(output, false);
        Assert.Equal(SpreadsheetDocumentType.MacroEnabledWorkbook, document.DocumentType);
        var vbaPart = document.WorkbookPart?.GetPartsOfType<VbaProjectPart>().SingleOrDefault();
        Assert.NotNull(vbaPart);
        using var embeddedStream = vbaPart!.GetStream();
        using var inputStream = File.OpenRead(vba);
        Assert.Equal(
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(inputStream)),
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(embeddedStream)));
    }
}

internal sealed class TestDirectory : IDisposable
{
    public string Root { get; }

    public TestDirectory(string? subdirectory = null)
    {
        Root = Path.Combine(WorkbookFixture.CreateDirectory(), subdirectory ?? "");
        Directory.CreateDirectory(Root);
    }

    public string PathFor(string name) => Path.Combine(Root, name);

    public void Dispose()
    {
        if (Directory.Exists(Root))
            Directory.Delete(Root, recursive: true);
    }
}
