using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Compression;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class ExcelModuleInsertionServiceTests
{
    [Fact]
    public async Task Coordinates_module_import_on_xlsm_copy_and_preserves_component_names_and_input_hash()
    {
        using var temp = new TestDirectory();
        var input = CreateSource(temp, "HR mẫu.xlsm");
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("Kết quả có VBA.xlsm");
        var sourceHash = WorkbookFixture.HashFile(input);
        var host = new FakeExcelAutomationHost(["ThisWorkbook", "Sheet1", "ModInHangLoat"]);

        var result = await new ExcelModuleInsertionService(host).InsertAsync(new(input, module, manifest, output));

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(input));
        Assert.Contains("ModInHangLoat", host.PersistedComponents);
        Assert.Contains("ProductSmokeTest", host.PersistedComponents);
        Assert.True(new ExcelPackageValidator().ValidateMacroWorkbook(output).IsValid);
        Assert.False(Directory.EnumerateFiles(temp.Root, ".tool-vba-*.tmp.xlsm").Any());
    }

    [Fact]
    public async Task Refuses_to_import_when_manifest_module_name_already_exists()
    {
        using var temp = new TestDirectory();
        var input = CreateSource(temp, "source.xlsm");
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ModInHangLoat");
        var output = temp.PathFor("result.xlsm");
        var host = new FakeExcelAutomationHost(["ThisWorkbook", "ModInHangLoat"]);

        var result = await new ExcelModuleInsertionService(host).InsertAsync(new(input, module, manifest, output));

        Assert.False(result.Succeeded);
        Assert.Contains("đã tồn tại", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(output));
        Assert.Equal(0, host.ImportCount);
        Assert.False(Directory.EnumerateFiles(temp.Root, ".tool-vba-*.tmp.xlsm").Any());
    }

    [Fact]
    public async Task Refuses_to_overwrite_existing_output_without_confirmation()
    {
        using var temp = new TestDirectory();
        var input = CreateSource(temp, "source.xlsm");
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("result.xlsm");
        var prior = new byte[] { 7, 8, 9 };
        File.WriteAllBytes(output, prior);
        var host = new FakeExcelAutomationHost(["ThisWorkbook"]);

        var result = await new ExcelModuleInsertionService(host).InsertAsync(new(input, module, manifest, output));

        Assert.False(result.Succeeded);
        Assert.Equal(prior, await File.ReadAllBytesAsync(output));
        Assert.Equal(0, host.OpenCount);
    }

    [Fact]
    public async Task Replaces_existing_output_only_after_confirmation_and_successful_validation()
    {
        using var temp = new TestDirectory();
        var input = CreateSource(temp, "source.xlsm");
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("result.xlsm");
        File.WriteAllBytes(output, [21, 22, 23]);
        var sourceHash = WorkbookFixture.HashFile(input);
        var host = new FakeExcelAutomationHost(["ThisWorkbook"]);

        var result = await new ExcelModuleInsertionService(host).InsertAsync(
            new ExcelModuleInsertionRequest(input, module, manifest, output, ConfirmOutputOverwrite: true));

        Assert.True(result.Succeeded, result.Message);
        Assert.NotEqual(WorkbookFixture.Hash([21, 22, 23]), WorkbookFixture.HashFile(output));
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(input));
        Assert.True(new ExcelPackageValidator().ValidateMacroWorkbook(output).IsValid);
    }

    [Fact]
    public async Task Keeps_input_and_existing_output_and_cleans_temporary_file_when_excel_import_fails()
    {
        using var temp = new TestDirectory();
        var input = CreateSource(temp, "source.xlsm");
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("result.xlsm");
        var prior = new byte[] { 11, 12, 13 };
        File.WriteAllBytes(output, prior);
        var sourceHash = WorkbookFixture.HashFile(input);
        var host = new FakeExcelAutomationHost(["ThisWorkbook"]) { ThrowOnImport = true };

        var result = await new ExcelModuleInsertionService(host).InsertAsync(
            new ExcelModuleInsertionRequest(input, module, manifest, output, ConfirmOutputOverwrite: true));

        Assert.False(result.Succeeded);
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(input));
        Assert.Equal(prior, await File.ReadAllBytesAsync(output));
        Assert.False(Directory.EnumerateFiles(temp.Root, ".tool-vba-*.tmp.xlsm").Any());
    }

    [Fact]
    public async Task Rejects_output_path_that_is_the_same_as_input()
    {
        using var temp = new TestDirectory();
        var input = CreateSource(temp, "same.xlsm");
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");

        var result = await new ExcelModuleInsertionService(new FakeExcelAutomationHost(["ThisWorkbook"]))
            .InsertAsync(new(input, module, manifest, input));

        Assert.False(result.Succeeded);
        Assert.Contains("trùng", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Imports_module_into_xlsx_copy_and_saves_macro_enabled_output()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("source.xlsx");
        WorkbookFixture.CreateWorkbook(input, SpreadsheetDocumentType.Workbook);
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("result.xlsm");
        var host = new FakeExcelAutomationHost(["ThisWorkbook"]);
        var sourceHash = WorkbookFixture.HashFile(input);

        var result = await new ExcelModuleInsertionService(host).InsertAsync(new(input, module, manifest, output));

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(input));
        Assert.Contains("ThisWorkbook", host.PersistedComponents);
        Assert.Contains("ProductSmokeTest", host.PersistedComponents);
        Assert.True(new ExcelPackageValidator().ValidateMacroWorkbook(output).IsValid);
        Assert.False(Directory.EnumerateFiles(temp.Root, ".tool-vba-*.tmp.*").Any());
    }

    [Fact]
    public async Task Accepts_xlsx_conversion_when_excel_changes_document_component_names()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("source.xlsx");
        WorkbookFixture.CreateWorkbook(input, SpreadsheetDocumentType.Workbook);
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("result.xlsm");
        var host = new FakeExcelAutomationHost(["ThisWorkbook", "Sheet1"])
        {
            VerificationComponentsOverride = ["ThisWorkbook", "Sheet2", "ProductSmokeTest", "ExtraDocument"]
        };

        var result = await new ExcelModuleInsertionService(host).InsertAsync(new(input, module, manifest, output));

        Assert.True(result.Succeeded, result.Message);
        Assert.True(File.Exists(output));
    }

    [Fact]
    public async Task Rejects_xlsx_conversion_when_saved_module_is_missing()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("source.xlsx");
        WorkbookFixture.CreateWorkbook(input, SpreadsheetDocumentType.Workbook);
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("result.xlsm");
        var host = new FakeExcelAutomationHost(["ThisWorkbook", "Sheet1"])
        {
            VerificationComponentsOverride = ["ThisWorkbook", "Sheet1"]
        };

        var result = await new ExcelModuleInsertionService(host).InsertAsync(new(input, module, manifest, output));

        Assert.False(result.Succeeded);
        Assert.Contains("module mới: thiếu", result.Message);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task Rejects_xlsm_conversion_when_original_vba_component_is_missing()
    {
        using var temp = new TestDirectory();
        var input = CreateSource(temp, "source.xlsm");
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("result.xlsm");
        var host = new FakeExcelAutomationHost(["ThisWorkbook", "OldModule"])
        {
            VerificationComponentsOverride = ["ThisWorkbook", "ProductSmokeTest"]
        };

        var result = await new ExcelModuleInsertionService(host).InsertAsync(new(input, module, manifest, output));

        Assert.False(result.Succeeded);
        Assert.Contains("component VBA gốc thiếu: 1", result.Message);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task Keeps_xlsx_source_and_existing_output_when_save_as_fails()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("source.xlsx");
        WorkbookFixture.CreateWorkbook(input, SpreadsheetDocumentType.Workbook);
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("result.xlsm");
        var prior = new byte[] { 41, 42, 43 };
        File.WriteAllBytes(output, prior);
        var sourceHash = WorkbookFixture.HashFile(input);
        var host = new FakeExcelAutomationHost(["ThisWorkbook"]) { ThrowOnSaveAs = true };

        var result = await new ExcelModuleInsertionService(host).InsertAsync(
            new ExcelModuleInsertionRequest(input, module, manifest, output, ConfirmOutputOverwrite: true));

        Assert.False(result.Succeeded);
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(input));
        Assert.Equal(prior, File.ReadAllBytes(output));
        Assert.False(Directory.EnumerateFiles(temp.Root, ".tool-vba-*.tmp.*").Any());
    }

    [Fact]
    public void Reports_locked_xlsx_as_in_use_instead_of_corrupt()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("source.xlsx");
        WorkbookFixture.CreateWorkbook(input, SpreadsheetDocumentType.Workbook);
        using var lockStream = new FileStream(input, FileMode.Open, FileAccess.ReadWrite, FileShare.None);

        var result = new ExcelVbaInjector().InspectWorkbook(input);

        Assert.False(result.IsValid);
        Assert.Contains("đang được", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Refuses_legacy_excel_macro_sheet_before_starting_excel_automation()
    {
        using var temp = new TestDirectory();
        var input = CreateSource(temp, "legacy-macrosheet.xlsm");
        using (var archive = ZipFile.Open(input, ZipArchiveMode.Update))
            archive.CreateEntry("xl/macrosheets/sheet1.xml");
        var module = WriteModule(temp, "new-module.bas");
        var manifest = WriteManifest(temp, module, "ProductSmokeTest");
        var output = temp.PathFor("result.xlsm");
        var host = new FakeExcelAutomationHost(["ThisWorkbook"]);

        var result = await new ExcelModuleInsertionService(host).InsertAsync(new(input, module, manifest, output));

        Assert.False(result.Succeeded);
        Assert.Contains("macro sheet", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, host.OpenCount);
        Assert.False(File.Exists(output));
    }

    private static string CreateSource(TestDirectory temp, string name)
    {
        var path = temp.PathFor(name);
        WorkbookFixture.CreateWorkbook(path, SpreadsheetDocumentType.MacroEnabledWorkbook, withVba: true);
        return path;
    }

    private static string WriteModule(TestDirectory temp, string name)
    {
        var path = temp.PathFor(name);
        File.WriteAllText(path, "Attribute VB_Name = \"ProductSmokeTest\"\r\nOption Explicit\r\n", new System.Text.UTF8Encoding(false));
        return path;
    }

    private static string WriteManifest(TestDirectory temp, string modulePath, string name)
    {
        using var stream = File.OpenRead(modulePath);
        var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
        var path = temp.PathFor("module.manifest.json");
        File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(new BasModuleManifest(name, hash)));
        return path;
    }
}

internal sealed class FakeExcelAutomationHost(IReadOnlyList<string> initialComponents) : IExcelAutomationHost
{
    private IReadOnlyList<string> _persisted = initialComponents.ToArray();
    public int ImportCount { get; private set; }
    public int OpenCount { get; private set; }
    public bool ThrowOnImport { get; init; }
    public bool ThrowOnSaveAs { get; init; }
    public IReadOnlyList<string>? VerificationComponentsOverride { get; init; }
    public IReadOnlyList<string> PersistedComponents => _persisted;

    public IExcelAutomationSession Open(string workbookPath, bool readOnly)
    {
        OpenCount++;
        return new Session(this, workbookPath,
            (OpenCount == 2 && VerificationComponentsOverride is not null
                ? VerificationComponentsOverride : _persisted).ToList());
    }

    private sealed class Session(FakeExcelAutomationHost owner, string workbookPath, List<string> components) : IExcelAutomationSession
    {
        public IReadOnlyList<string> GetComponentNames() => components.ToArray();

        public string ImportStandardModule(string modulePath)
        {
            if (owner.ThrowOnImport) throw new ExcelAutomationException("Excel không thể import module vào bản tạm.");
            owner.ImportCount++;
            components.Add("ProductSmokeTest");
            return "ProductSmokeTest";
        }

        public void Save() => owner._persisted = components.ToArray();
        public void SaveAsMacroEnabled(string outputPath)
        {
            if (owner.ThrowOnSaveAs) throw new ExcelAutomationException("Excel không lưu được workbook tạm dạng .xlsm.");
            File.Copy(workbookPath, outputPath);
            using var document = SpreadsheetDocument.Open(outputPath, true);
            document.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);
            var part = document.WorkbookPart!.AddNewPart<VbaProjectPart>();
            using var stream = new MemoryStream(WorkbookFixture.VbaBytes);
            part.FeedData(stream);
            owner._persisted = components.ToArray();
        }
        public void Dispose() { }
    }
}
