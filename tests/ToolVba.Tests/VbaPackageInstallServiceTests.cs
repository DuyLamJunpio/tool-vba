using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class VbaPackageInstallServiceTests : IDisposable
{
    private static readonly LicenseProduct Product = new("test-01", "1.0.0");
    private readonly TestDirectory _temp = new();

    public void Dispose() => _temp.Dispose();

    [Fact]
    public async Task Installs_package_into_xlsx_copy_and_saves_macro_enabled_output()
    {
        var input = CreateXlsx("Báo cáo.xlsx");
        var output = _temp.PathFor("test-01_1.0.0_ban-quyen.xlsm");
        var sourceHash = WorkbookFixture.HashFile(input);
        var host = new FakeVbaWorkbookHost(Documents());

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(input, output));

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(input));
        Assert.Equal(["ThisWorkbook", "Sheet1", "modReport", "clsRow", "modLicenseApiV1", "frmInput"], host.PersistedNames);
        Assert.True(DocumentModuleMerger.ContainsManagedCode(host.Persisted["ThisWorkbook"].Code));
        Assert.Equal("2026.11.7.1", host.PersistedProperties["ToolVba.PackageVersion"]);
        Assert.Equal("modReport;clsRow;modLicenseApiV1;frmInput", host.PersistedProperties["ToolVba.Managed.1"]);
        Assert.True(new ExcelPackageValidator().ValidateMacroWorkbook(output).IsValid);
        AssertNoTemporaryFiles();
    }

    [Fact]
    public async Task Fills_the_licensing_module_for_the_product_and_every_other_sheet()
    {
        var host = new FakeVbaWorkbookHost(Documents(), [new("Dữ liệu", false, SheetVisibility.Visible), new("Báo cáo", false, SheetVisibility.Hidden)]);

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(CreateXlsx("a.xlsx"), _temp.PathFor("out.xlsm")));

        Assert.True(result.Succeeded, result.Message);
        var module = host.Persisted["modLicenseApiV1"].Code;
        Assert.Contains("Cfg_ProductKey = \"test-01\"", module);
        Assert.Contains("Cfg_ProductVersion = \"1.0.0\"", module);
        Assert.Contains("protectedNames(0) = \"D\" & ChrW(7919) & \" li\" & ChrW(7879) & \"u\"", module);
        Assert.Contains("protectedNames(1) = \"B\" & ChrW(225) & \"o c\" & ChrW(225) & \"o\"", module);
    }

    [Fact]
    public async Task Adds_the_guide_sheet_first_and_saves_every_other_sheet_very_hidden()
    {
        var host = new FakeVbaWorkbookHost(Documents(), [new("Dữ liệu", false, SheetVisibility.Visible), new("Báo cáo", false, SheetVisibility.Visible)]);

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(CreateXlsx("a.xlsx"), _temp.PathFor("out.xlsm")));

        Assert.True(result.Succeeded, result.Message);
        Assert.Equal(
            [new("Huong dan", false, SheetVisibility.Visible), new("Dữ liệu", false, SheetVisibility.VeryHidden), new("Báo cáo", false, SheetVisibility.VeryHidden)],
            host.PersistedSheets);
        Assert.Equal(GuideSheet.Lines, host.GuideLines);
        Assert.Equal("Huong dan", host.PersistedProperties[GuideSheet.MarkerProperty]);
    }

    [Fact]
    public async Task Packaging_its_own_output_again_reuses_the_guide_sheet()
    {
        var host = new FakeVbaWorkbookHost(Documents(), [new("Dữ liệu", false, SheetVisibility.Visible)]);
        var service = new VbaPackageInstallService(host);
        Assert.True((await service.InstallAsync(Request(CreateXlsx("a.xlsx"), _temp.PathFor("first.xlsm")))).Succeeded);

        var second = await service.InstallAsync(Request(CreateXlsm("first-copy.xlsm"), _temp.PathFor("second.xlsm")));

        Assert.True(second.Succeeded, second.Message);
        Assert.Equal(2, host.PersistedSheets.Count);
        Assert.Single(host.PersistedSheets, s => s.Name == "Huong dan");
    }

    [Fact]
    public async Task Refuses_a_users_own_guide_named_sheet_even_when_an_old_marker_remains()
    {
        // Packaged once, then the author deleted the tool's guide and made their own "Huong dan".
        var host = new FakeVbaWorkbookHost(Documents(), [new("Huong dan", false, SheetVisibility.Visible), new("Dữ liệu", false, SheetVisibility.Visible)]);
        host.PersistedProperties[GuideSheet.MarkerProperty] = GuideSheet.Name;
        host.CellA1["Huong dan"] = "Ghi chú riêng của tác giả";

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(CreateXlsm("a.xlsm"), _temp.PathFor("out.xlsm")));

        Assert.False(result.Succeeded);
        Assert.Contains("Huong dan", result.Message);
        Assert.Equal(0, host.SaveCount);
    }

    [Fact]
    public async Task Stops_when_the_product_has_its_own_Auto_Open()
    {
        var host = new FakeVbaWorkbookHost([
            .. Documents(),
            new("modUser", VbaComponentKind.Standard, "Attribute VB_Name = \"modUser\"\r\nSub Auto_Open()\r\nEnd Sub\r\n")
        ]);

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(CreateXlsm("a.xlsm"), _temp.PathFor("out.xlsm")));

        Assert.False(result.Succeeded);
        Assert.Contains("modUser", result.Message);
        Assert.Contains("Auto_Open", result.Message);
        Assert.Equal(0, host.SaveCount);
    }

    [Fact]
    public async Task Stops_when_the_workbook_structure_is_protected()
    {
        var host = new FakeVbaWorkbookHost(Documents()) { StructureProtected = true };

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(CreateXlsx("a.xlsx"), _temp.PathFor("out.xlsm")));

        Assert.False(result.Succeeded);
        Assert.Contains("cấu trúc", result.Message);
        Assert.Equal(0, host.SaveCount);
    }

    [Fact]
    public async Task Stops_before_excel_when_the_package_has_no_licensing_module()
    {
        var host = new FakeVbaWorkbookHost(Documents());

        var result = await new VbaPackageInstallService(host).InstallAsync(
            new(CreateXlsx("a.xlsx"), _temp.PathFor("out.xlsm"), ReadPackage(includeLicense: false), Product));

        Assert.False(result.Succeeded);
        Assert.Contains("module bản quyền", result.Message);
        Assert.Equal(0, host.OpenCount);
    }

    [Fact]
    public async Task Gives_the_handler_advice_only_and_not_the_rename_advice_for_an_unusual_signature()
    {
        var host = new FakeVbaWorkbookHost([
            new("ThisWorkbook", VbaComponentKind.Workbook, "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Huy As Boolean)\r\nEnd Sub\r\n"),
            new("Sheet1", VbaComponentKind.Sheet, "")
        ]);

        var result = await new VbaPackageInstallService(host).InstallAsync(
            new(CreateXlsm("a.xlsm"), _temp.PathFor("out.xlsm"), ReadPackageWithWorkbookCode(LicenseEvents), Product));

        Assert.False(result.Succeeded);
        Assert.Contains("chữ ký", result.Message);
        Assert.DoesNotContain("đổi tên hoặc xóa phần code trùng", result.Message);
    }

    [Fact]
    public async Task Replaces_managed_modules_in_xlsm_and_keeps_user_modules()
    {
        var input = CreateXlsm("source.xlsm");
        var output = _temp.PathFor("result.xlsm");
        var host = new FakeVbaWorkbookHost([
            .. Documents(),
            new("modReport", VbaComponentKind.Standard, "Attribute VB_Name = \"modReport\"\r\n' old version"),
            new("modUser", VbaComponentKind.Standard, "Attribute VB_Name = \"modUser\"\r\n' employee code")
        ]);

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(input, output));

        Assert.True(result.Succeeded, result.Message);
        Assert.Contains("modUser", host.PersistedNames);
        Assert.DoesNotContain("modReport1", host.PersistedNames);
        Assert.Contains("Public Sub Prepare()", host.Persisted["modReport"].Code);
        Assert.Equal("' employee code", host.Persisted["modUser"].Code.Split("\r\n")[1]);
    }

    [Fact]
    public async Task Removes_module_managed_by_an_older_package_that_is_no_longer_shipped()
    {
        var input = CreateXlsm("source.xlsm");
        var output = _temp.PathFor("result.xlsm");
        var host = new FakeVbaWorkbookHost([
            .. Documents(),
            new("modRetired", VbaComponentKind.Standard, "Attribute VB_Name = \"modRetired\"\r\n")
        ]);
        host.PersistedProperties["ToolVba.Managed.1"] = "modReport;modRetired";

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(input, output));

        Assert.True(result.Succeeded, result.Message);
        Assert.DoesNotContain("modRetired", host.PersistedNames);
    }

    [Fact]
    public async Task Stops_without_output_when_workbook_already_has_the_same_event_procedure()
    {
        var input = CreateXlsm("source.xlsm");
        var output = _temp.PathFor("result.xlsm");
        var host = new FakeVbaWorkbookHost([
            new("ThisWorkbook", VbaComponentKind.Workbook, "Private Sub Workbook_Open()\r\n    MsgBox \"mine\"\r\nEnd Sub\r\n"),
            new("Sheet1", VbaComponentKind.Sheet, "")
        ]);
        var sourceHash = WorkbookFixture.HashFile(input);

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(input, output));

        Assert.False(result.Succeeded);
        Assert.Contains("Workbook_Open", result.Message);
        Assert.False(File.Exists(output));
        Assert.Equal(0, host.SaveCount);
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(input));
        AssertNoTemporaryFiles();
    }

    [Fact]
    public async Task Stops_when_package_contains_sheet_code_for_a_sheet_the_workbook_lacks()
    {
        var input = CreateXlsx("source.xlsx");
        var output = _temp.PathFor("result.xlsm");
        var host = new FakeVbaWorkbookHost(Documents());

        var result = await new VbaPackageInstallService(host).InstallAsync(new(input, output, ReadPackageWithSheet("Sheet9"), Product));

        Assert.False(result.Succeeded);
        Assert.Contains("Sheet9", result.Message);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task Fills_existing_sheet_module_by_code_name()
    {
        var input = CreateXlsx("source.xlsx");
        var output = _temp.PathFor("result.xlsm");
        var host = new FakeVbaWorkbookHost(Documents());

        var result = await new VbaPackageInstallService(host).InstallAsync(new(input, output, ReadPackageWithSheet("Sheet1"), Product));

        Assert.True(result.Succeeded, result.Message);
        Assert.Contains("Worksheet_Change", host.Persisted["Sheet1"].Code);
    }

    [Fact]
    public async Task Matches_sheet_by_tab_name_when_code_name_is_localized()
    {
        var input = CreateXlsx("source.xlsx");
        var output = _temp.PathFor("result.xlsm");
        var host = new FakeVbaWorkbookHost([
            new("SổlàmviệcNày", VbaComponentKind.Workbook, ""),
            new("Trang_tính1", VbaComponentKind.Sheet, "", SheetName: "Sheet9")
        ]);

        var result = await new VbaPackageInstallService(host).InstallAsync(new(input, output, ReadPackageWithSheet("Sheet9", sheetName: "Sheet9"), Product));

        Assert.True(result.Succeeded, result.Message);
        Assert.Contains("Worksheet_Change", host.Persisted["Trang_tính1"].Code);
    }

    [Fact]
    public async Task Renames_localized_workbook_module_to_ascii_before_pasting_code()
    {
        var input = CreateXlsx("source.xlsx");
        var output = _temp.PathFor("result.xlsm");
        var host = new FakeVbaWorkbookHost([
            new("SổlàmviệcNày", VbaComponentKind.Workbook, ""),
            new("Trang_tính1", VbaComponentKind.Sheet, "", SheetName: "Sheet1")
        ]);

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(input, output));

        Assert.True(result.Succeeded, result.Message);
        Assert.False(host.Persisted.ContainsKey("SổlàmviệcNày"));
        Assert.True(DocumentModuleMerger.ContainsManagedCode(host.Persisted["ThisWorkbook"].Code));
    }

    [Fact]
    public async Task Stops_before_touching_a_workbook_module_whose_name_excel_lost()
    {
        var input = CreateXlsm("source.xlsm");
        var output = _temp.PathFor("result.xlsm");
        var host = new FakeVbaWorkbookHost([new("", VbaComponentKind.Workbook, ""), new("Sheet1", VbaComponentKind.Sheet, "")]);

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(input, output));

        Assert.False(result.Succeeded);
        Assert.Contains("ThisWorkbook", result.Message);
        Assert.Equal(0, host.SaveCount);
        Assert.Equal(0, host.DocumentReads);
    }

    [Fact]
    public async Task Refuses_existing_output_without_confirmation_before_starting_excel()
    {
        var input = CreateXlsm("source.xlsm");
        var output = _temp.PathFor("result.xlsm");
        File.WriteAllBytes(output, [1, 2, 3]);
        var host = new FakeVbaWorkbookHost(Documents());

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(input, output));

        Assert.False(result.Succeeded);
        Assert.Equal([1, 2, 3], File.ReadAllBytes(output));
        Assert.Equal(0, host.OpenCount);
    }

    [Fact]
    public async Task Rejects_output_path_equal_to_input()
    {
        var input = CreateXlsm("same.xlsm");

        var result = await new VbaPackageInstallService(new FakeVbaWorkbookHost(Documents()))
            .InstallAsync(new(input, input, ReadPackage(), Product, ConfirmOutputOverwrite: true));

        Assert.False(result.Succeeded);
        Assert.Contains("trùng", result.Message);
    }

    [Fact]
    public async Task Keeps_input_and_existing_output_when_excel_import_fails()
    {
        var input = CreateXlsm("source.xlsm");
        var output = _temp.PathFor("result.xlsm");
        File.WriteAllBytes(output, [9, 9]);
        var sourceHash = WorkbookFixture.HashFile(input);
        var host = new FakeVbaWorkbookHost(Documents()) { ThrowOnImport = true };

        var result = await new VbaPackageInstallService(host)
            .InstallAsync(new(input, output, ReadPackage(), Product, ConfirmOutputOverwrite: true));

        Assert.False(result.Succeeded);
        Assert.Equal([9, 9], File.ReadAllBytes(output));
        Assert.Equal(sourceHash, WorkbookFixture.HashFile(input));
        AssertNoTemporaryFiles();
    }

    [Fact]
    public async Task Fails_when_excel_renames_imported_component()
    {
        var input = CreateXlsx("source.xlsx");
        var output = _temp.PathFor("result.xlsm");
        var host = new FakeVbaWorkbookHost(Documents()) { RenameImportedTo = "modReport1" };

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(input, output));

        Assert.False(result.Succeeded);
        Assert.Contains("modReport", result.Message);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public async Task Refuses_package_files_changed_after_they_were_verified()
    {
        var input = CreateXlsx("source.xlsx");
        var output = _temp.PathFor("result.xlsm");
        var package = ReadPackage();
        File.AppendAllText(package.Components[0].FilePath, "' changed");
        var host = new FakeVbaWorkbookHost(Documents());

        var result = await new VbaPackageInstallService(host).InstallAsync(new(input, output, package, Product));

        Assert.False(result.Succeeded);
        Assert.Equal(0, host.OpenCount);
    }

    [Fact]
    public async Task Refuses_legacy_macro_sheet_before_starting_excel()
    {
        var input = CreateXlsm("legacy.xlsm");
        using (var archive = ZipFile.Open(input, ZipArchiveMode.Update))
            archive.CreateEntry("xl/macrosheets/sheet1.xml");
        var host = new FakeVbaWorkbookHost(Documents());

        var result = await new VbaPackageInstallService(host).InstallAsync(Request(input, _temp.PathFor("out.xlsm")));

        Assert.False(result.Succeeded);
        Assert.Equal(0, host.OpenCount);
    }

    [Fact]
    public void Managed_list_is_split_into_property_sized_chunks_and_read_back()
    {
        var names = Enumerable.Range(1, 30).Select(i => $"modComponentNumber{i:00}").ToArray();

        var chunks = ManagedComponentList.ToChunks(names);

        Assert.True(chunks.Count > 1);
        Assert.All(chunks, chunk => Assert.True(chunk.Length <= ManagedComponentList.MaxChunkLength));
        Assert.Equal(names, ManagedComponentList.FromChunks(chunks));
    }

    private const string LicenseEvents =
        "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\r\n" +
        "    If Not modLicenseApiV1.License_BeforeSave() Then Cancel = True: Exit Sub\r\n" +
        "End Sub\r\n";

    private VbaPackageInstallRequest Request(string input, string output) => new(input, output, ReadPackage(), Product);

    private static FakeComponent[] Documents() =>
    [
        new("ThisWorkbook", VbaComponentKind.Workbook, ""),
        new("Sheet1", VbaComponentKind.Sheet, "")
    ];

    private string CreateXlsx(string name)
    {
        var path = _temp.PathFor(name);
        WorkbookFixture.CreateWorkbook(path, SpreadsheetDocumentType.Workbook);
        return path;
    }

    private string CreateXlsm(string name)
    {
        var path = _temp.PathFor(name);
        WorkbookFixture.CreateWorkbook(path, SpreadsheetDocumentType.MacroEnabledWorkbook, withVba: true);
        return path;
    }

    private VbaPackage ReadPackage(bool includeLicense = true)
    {
        var zip = PackageFixture.BuildPackage(_temp, $"pkg-{Guid.NewGuid():N}.zip", includeForm: true, includeLicense: includeLicense);
        return Read(zip);
    }

    private VbaPackage ReadPackageWithSheet(string codeName, string? sheetName = null)
    {
        var staging = Directory.CreateDirectory(_temp.PathFor("sheet-" + Guid.NewGuid().ToString("N"))).FullName;
        var sheet = Path.Combine(staging, "SheetCode.txt");
        File.WriteAllText(sheet, "Private Sub Worksheet_Change(ByVal Target As Range)\r\nEnd Sub\r\n", new UTF8Encoding(false));
        return BuildWith(staging, new(codeName, VbaComponentKind.Sheet, sheet, SheetName: sheetName));
    }

    private VbaPackage ReadPackageWithWorkbookCode(string code)
    {
        var staging = Directory.CreateDirectory(_temp.PathFor("wb-" + Guid.NewGuid().ToString("N"))).FullName;
        var workbook = Path.Combine(staging, "ThisWorkbook.txt");
        File.WriteAllText(workbook, code, new UTF8Encoding(false));
        return BuildWith(staging, new("ThisWorkbook", VbaComponentKind.Workbook, workbook));
    }

    private VbaPackage BuildWith(string staging, VbaPackageSource document)
    {
        var module = PackageFixture.WriteAnsi(staging, "modReport.bas", PackageFixture.StandardSource("modReport"));
        var license = PackageFixture.WriteAnsi(staging, "modLicenseApiV1.bas", PackageFixture.LicenseModuleSource);
        var zip = _temp.PathFor($"doc-{Guid.NewGuid():N}.zip");
        Assert.True(VbaPackageBuilder.Build(
        [
            new("modReport", VbaComponentKind.Standard, module),
            new(PackageFixture.LicenseModuleName, VbaComponentKind.Standard, license),
            document
        ], zip).IsValid);
        return Read(zip);
    }

    private VbaPackage Read(string zip)
    {
        var result = new VbaPackageReader().Read(zip, "2026.11.7.1", _temp.PathFor("pkg-" + Guid.NewGuid().ToString("N")));
        Assert.True(result.IsValid, result.Error);
        return result.Package!;
    }

    private void AssertNoTemporaryFiles() =>
        Assert.False(Directory.EnumerateFiles(_temp.Root, ".tool-vba-*", SearchOption.TopDirectoryOnly).Any());
}

internal sealed record FakeComponent(string Name, VbaComponentKind Kind, string Code, string? SheetName = null);

internal sealed class FakeVbaWorkbookHost(IEnumerable<FakeComponent> initial, IEnumerable<WorkbookSheetInfo>? sheets = null) : IVbaWorkbookHost
{
    public Dictionary<string, FakeComponent> Persisted { get; private set; } =
        initial.ToDictionary(c => c.Name, StringComparer.OrdinalIgnoreCase);
    public List<string> PersistedNames => Persisted.Values.Select(c => c.Name).ToList();
    public Dictionary<string, string> PersistedProperties { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<WorkbookSheetInfo> PersistedSheets { get; private set; } =
        sheets?.ToList() ?? [new("Sheet1", false, SheetVisibility.Visible)];
    public IReadOnlyList<string> GuideLines { get; private set; } = [];
    /// <summary>Cell A1 of sheets the test sets up; the guide sheet's A1 comes from its written lines.</summary>
    public Dictionary<string, string> CellA1 { get; } = new(StringComparer.OrdinalIgnoreCase);
    public int OpenCount { get; private set; }
    public int SaveCount { get; private set; }
    public int DocumentReads { get; private set; }
    public bool ThrowOnImport { get; init; }
    public bool StructureProtected { get; init; }
    public string? RenameImportedTo { get; init; }

    public IVbaWorkbookSession Open(string workbookPath, bool readOnly)
    {
        OpenCount++;
        return new Session(this, workbookPath);
    }

    private sealed class Session(FakeVbaWorkbookHost owner, string workbookPath) : IVbaWorkbookSession
    {
        private readonly Dictionary<string, FakeComponent> _components = new(owner.Persisted, StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _properties = new(owner.PersistedProperties, StringComparer.OrdinalIgnoreCase);
        private readonly List<WorkbookSheetInfo> _sheets = [.. owner.PersistedSheets];
        private IReadOnlyList<string> _guideLines = owner.GuideLines;

        public IReadOnlyList<VbaComponentInfo> GetComponents() =>
            _components.Values.Select(c => new VbaComponentInfo(c.Name, c.Kind, c.SheetName)).ToList();

        public string ImportComponent(string filePath)
        {
            if (owner.ThrowOnImport) throw new ExcelAutomationException("Excel không thể import component vào bản tạm.");
            var code = File.ReadAllText(filePath, Encoding.Latin1);
            var name = code.Split("\r\n").First(l => l.StartsWith("Attribute VB_Name")).Split('"')[1];
            if (_components.ContainsKey(name)) name += "1";
            if (owner.RenameImportedTo is not null && name == "modReport") name = owner.RenameImportedTo;
            var kind = Path.GetExtension(filePath).ToLowerInvariant() switch
            {
                ".cls" => VbaComponentKind.Class,
                ".frm" => VbaComponentKind.Form,
                _ => VbaComponentKind.Standard
            };
            if (kind == VbaComponentKind.Form && !File.Exists(Path.ChangeExtension(filePath, ".frx")))
                throw new ExcelAutomationException("Thiếu .frx");
            _components[name] = new(name, kind, code);
            return name;
        }

        public void RemoveComponent(string name)
        {
            if (_components[name].Kind is VbaComponentKind.Workbook or VbaComponentKind.Sheet)
                throw new InvalidOperationException("Document modules cannot be removed.");
            _components.Remove(name);
        }

        public string GetDocumentCode(string name)
        {
            // Real Excel crashes when the code module of a workbook module with a lost name is read.
            if (name.Length == 0) throw new InvalidOperationException("Excel would crash here.");
            owner.DocumentReads++;
            return _components[name].Code;
        }
        public void RenameDocumentModule(string name, string newName)
        {
            var component = _components[name];
            _components.Remove(name);
            _components[newName] = component with { Name = newName };
        }

        public void SetDocumentCode(string name, string code) => _components[name] = _components[name] with { Code = code };
        public void ExportComponent(string name, string filePath)
        {
            File.WriteAllText(filePath, _components[name].Code, Encoding.Latin1);
            if (_components[name].Kind == VbaComponentKind.Form)
                File.WriteAllBytes(Path.ChangeExtension(filePath, ".frx"), [5, 6, 7]);
        }
        public string? GetCustomProperty(string name) => _properties.TryGetValue(name, out var value) ? value : null;
        public void SetCustomProperty(string name, string value) => _properties[name] = value;
        public void DeleteCustomProperty(string name) => _properties.Remove(name);

        public IReadOnlyList<WorkbookSheetInfo> GetSheets() => [.. _sheets];
        public bool IsStructureProtected() => owner.StructureProtected;

        public string ReadCell(string sheetName, string address) =>
            owner.CellA1.TryGetValue(sheetName, out var value) ? value
            : string.Equals(sheetName, GuideSheet.Name, StringComparison.OrdinalIgnoreCase) && _guideLines.Count > 0 ? _guideLines[0]
            : "";

        public void WriteGuideSheet(string name, IReadOnlyList<string> lines, bool create)
        {
            var existing = _sheets.FindIndex(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
            if (create == existing >= 0) throw new InvalidOperationException("Guide sheet already exists or is missing.");
            if (existing >= 0) _sheets.RemoveAt(existing);
            _sheets.Insert(0, new(name, false, SheetVisibility.Visible));
            _guideLines = [.. lines];
        }

        public void SetSheetVisibility(string name, SheetVisibility visibility)
        {
            var index = _sheets.FindIndex(s => s.Name == name);
            _sheets[index] = _sheets[index] with { Visibility = visibility };
        }

        public void Save() => Commit();

        public void SaveAsMacroEnabled(string outputPath)
        {
            File.Copy(workbookPath, outputPath);
            using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(outputPath, true))
            {
                document.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);
                var part = document.WorkbookPart!.AddNewPart<DocumentFormat.OpenXml.Packaging.VbaProjectPart>();
                using var stream = new MemoryStream(WorkbookFixture.VbaBytes);
                part.FeedData(stream);
            }
            Commit();
        }

        private void Commit()
        {
            owner.SaveCount++;
            owner.Persisted = new(_components, StringComparer.OrdinalIgnoreCase);
            owner.PersistedProperties = new(_properties, StringComparer.OrdinalIgnoreCase);
            owner.PersistedSheets = [.. _sheets];
            owner.GuideLines = _guideLines;
        }

        public void Dispose() { }
    }
}
