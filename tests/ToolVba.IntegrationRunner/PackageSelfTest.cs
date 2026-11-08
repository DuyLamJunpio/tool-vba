using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ToolVba.Core;

/// <summary>
/// End-to-end check against real Excel with synthetic workbooks and the real licensing module
/// (license-sample/api-v1-one-file): package v1 into an .xlsx (guide sheet, hidden sheets, filled module),
/// add an employee module, update to v2, merge the licensing call into an existing BeforeSave handler,
/// and confirm a conflicting Workbook_Open and a product Auto_Open are refused. Needs Excel and VBA access.
/// </summary>
internal static class PackageSelfTest
{
    private static readonly Encoding Ansi = Encoding.Latin1;
    private static readonly LicenseProduct Product = new("selftest", "1.0.0");
    private static int _failures;

    public static async Task<int> RunAsync(string directoryArg)
    {
        var root = Path.GetFullPath(directoryArg);
        if (Directory.Exists(root) && Directory.EnumerateFileSystemEntries(root).Any())
        {
            Console.Error.WriteLine("FAIL: Thư mục self-test phải trống hoặc chưa tồn tại.");
            return 2;
        }
        Directory.CreateDirectory(root);
        IVbaWorkbookHost workbooks = new ExcelAutomationHost();
        var installer = new VbaPackageInstallService(workbooks);

        var source = CreateWorkbook(Path.Combine(root, "source.xlsx"));
        var v1 = BuildPackage(root, "v1", "2026.1.1.1", includeClass: true);
        var out1 = Path.Combine(root, "out1.xlsm");
        if (!Check("Đóng gói v1 vào .xlsx", await installer.InstallAsync(new(source, out1, v1, Product)))) return Finish(root);
        VerifyLicensing(workbooks, out1);

        AddModule(workbooks, out1, root, "modUser", "Public Sub Mine()\r\nEnd Sub\r\n");
        var v2 = BuildPackage(root, "v2", "2026.1.1.2", includeClass: false);
        var out2 = Path.Combine(root, "out2.xlsm");
        if (Check("Cập nhật lên v2 trên file đã đóng gói", await installer.InstallAsync(new(out1, out2, v2, Product))))
        {
            VerifyUpdate(workbooks, out2);
            VerifyLicensing(workbooks, out2);
        }

        var conflicted = CopyWithWorkbookCode(workbooks, out1, root, "conflict.xlsm",
            "Private Sub Workbook_Open()\r\n    MsgBox \"employee\"\r\nEnd Sub\r\n");
        var conflict = await installer.InstallAsync(new(conflicted, Path.Combine(root, "out3.xlsm"), v1, Product));
        Expect("Từ chối Workbook_Open trùng", !conflict.Succeeded && conflict.Message.Contains("Workbook_Open"), conflict.Message);

        var hooked = CopyWithWorkbookCode(workbooks, out1, root, "hook.xlsm",
            "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\r\n    Application.StatusBar = \"employee\"\r\nEnd Sub\r\n");
        var out4 = Path.Combine(root, "out4.xlsm");
        if (Check("Ghép sự kiện bản quyền vào BeforeSave có sẵn", await installer.InstallAsync(new(hooked, out4, v1, Product))))
            VerifyHook(workbooks, out4);

        var autoOpen = Path.Combine(root, "auto-open.xlsm");
        File.Copy(out1, autoOpen);
        AddModule(workbooks, autoOpen, root, "modStart", "Public Sub Auto_Open()\r\nEnd Sub\r\n");
        var refused = await installer.InstallAsync(new(autoOpen, Path.Combine(root, "out5.xlsm"), v1, Product));
        Expect("Từ chối Auto_Open riêng của sản phẩm", !refused.Succeeded && refused.Message.Contains("Auto_Open"), refused.Message);

        return Finish(root);
    }

    private static int Finish(string root)
    {
        Console.WriteLine(_failures == 0 ? "PASS: Self-test hoàn tất." : $"FAIL: {_failures} bước không đạt.");
        Console.WriteLine("Thư mục: " + root);
        return _failures == 0 ? 0 : 1;
    }

    private static void VerifyLicensing(IVbaWorkbookHost workbooks, string path)
    {
        using var session = workbooks.Open(path, readOnly: true);
        var sheets = session.GetSheets();
        var detail = string.Join(", ", sheets.Select(s => $"{s.Name}={s.Visibility}"));
        Expect("Sheet đầu là \"Huong dan\" và đang hiện",
            sheets.Count > 0 && sheets[0].Name == GuideSheet.Name && sheets[0].Visibility == SheetVisibility.Visible, detail);
        Expect("Có đúng một sheet hướng dẫn", sheets.Count(s => s.Name == GuideSheet.Name) == 1, detail);
        Expect("Mọi sheet khác very hidden", sheets.Skip(1).All(s => s.Visibility == SheetVisibility.VeryHidden), detail);
        var module = session.GetDocumentCode("modLicenseApiV1");
        Expect("Module bản quyền có mã sản phẩm và phiên bản",
            module.Contains("Cfg_ProductKey = \"selftest\"") && module.Contains("Cfg_ProductVersion = \"1.0.0\""), "(không thấy)");
        Expect("Tên sheet có dấu ghi bằng ChrW", module.Contains("\"D\" & ChrW(7919) & \" li\" & ChrW(7879) & \"u\""), "(không thấy)");
    }

    private static void VerifyUpdate(IVbaWorkbookHost workbooks, string path)
    {
        using var session = workbooks.Open(path, readOnly: true);
        var names = session.GetComponents().Select(c => c.Name).ToList();
        Expect("Không có module tài liệu bị mất tên", names.All(n => n.Length > 0), string.Join(", ", names));
        Expect("Module workbook tên ThisWorkbook", session.GetComponents().Any(c => c.Kind == VbaComponentKind.Workbook && c.Name == "ThisWorkbook"), string.Join(", ", names));
        Expect("Giữ module tự viết của nhân viên", names.Contains("modUser"), string.Join(", ", names));
        Expect("Gỡ module không còn trong gói", !names.Contains("clsRow"), string.Join(", ", names));
        Expect("Không có tên bị đổi kiểu modReport1", !names.Any(n => n.StartsWith("modReport", StringComparison.Ordinal) && n != "modReport"), string.Join(", ", names));
        var workbookCode = session.GetDocumentCode(WorkbookModule(session));
        Expect("ThisWorkbook có đúng một vùng code quản lý",
            workbookCode.Split("\r\n").Count(l => l.Trim() == DocumentModuleMerger.CodeBegin) == 1, workbookCode);
        Expect("Ghi phiên bản v2", session.GetCustomProperty(ManagedComponentList.VersionProperty) == "2026.1.1.2",
            session.GetCustomProperty(ManagedComponentList.VersionProperty) ?? "(trống)");
    }

    private static void VerifyHook(IVbaWorkbookHost workbooks, string path)
    {
        using var session = workbooks.Open(path, readOnly: true);
        var lines = session.GetDocumentCode(WorkbookModule(session)).Split("\r\n").Select(l => l.Trim()).ToList();
        var signature = lines.FindIndex(l => l.StartsWith("Private Sub Workbook_BeforeSave", StringComparison.Ordinal));
        Expect("Có đúng một Workbook_BeforeSave", lines.Count(l => l.StartsWith("Private Sub Workbook_BeforeSave", StringComparison.Ordinal)) == 1, string.Join(" | ", lines));
        Expect("Dòng bản quyền nằm đầu thủ tục, code nhân viên giữ nguyên",
            signature >= 0 && lines[signature + 1] == DocumentModuleMerger.HookBegin &&
            lines[signature + 2].StartsWith("If Not modLicenseApiV1.License_BeforeSave() Then Cancel = True: Exit Sub", StringComparison.Ordinal) &&
            lines.Contains("Application.StatusBar = \"employee\""), string.Join(" | ", lines));
    }

    private static string CopyWithWorkbookCode(IVbaWorkbookHost workbooks, string from, string root, string name, string code)
    {
        var path = Path.Combine(root, name);
        File.Copy(from, path);
        using var session = workbooks.Open(path, readOnly: false);
        session.SetDocumentCode(WorkbookModule(session), code);
        session.Save();
        return path;
    }

    private static void AddModule(IVbaWorkbookHost workbooks, string path, string root, string module, string body)
    {
        var file = Path.Combine(root, module + ".bas");
        File.WriteAllBytes(file, Ansi.GetBytes($"Attribute VB_Name = \"{module}\"\r\nOption Explicit\r\n{body}"));
        using var session = workbooks.Open(path, readOnly: false);
        session.ImportComponent(file);
        session.Save();
    }

    private static string WorkbookModule(IVbaWorkbookSession session) =>
        session.GetComponents().First(c => c.Kind == VbaComponentKind.Workbook).Name;

    private static string LicenseSample(string file)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "ToolVba.sln")))
            directory = directory.Parent;
        return Path.Combine(directory?.FullName ?? throw new InvalidOperationException("Không tìm thấy ToolVba.sln."),
            "license-sample", "api-v1-one-file", file);
    }

    private static VbaPackage BuildPackage(string root, string label, string version, bool includeClass)
    {
        var staging = Directory.CreateDirectory(Path.Combine(root, "staging-" + label)).FullName;
        string WriteAnsi(string name, string text) { var p = Path.Combine(staging, name); File.WriteAllBytes(p, Ansi.GetBytes(text)); return p; }
        string WriteUtf8(string name, string text) { var p = Path.Combine(staging, name); File.WriteAllText(p, text, new UTF8Encoding(false)); return p; }
        var events = string.Join("\r\n", File.ReadAllLines(LicenseSample("ThisWorkbook-events.txt")).Where(l => !l.TrimStart().StartsWith('\'')));
        var license = Path.Combine(staging, "modLicenseApiV1.bas");
        File.Copy(LicenseSample("modLicenseApiV1.bas"), license);

        var sources = new List<VbaPackageSource>
        {
            new("modReport", VbaComponentKind.Standard, WriteAnsi("modReport.bas",
                $"Attribute VB_Name = \"modReport\"\r\nOption Explicit\r\nPublic Function PackageLabel() As String\r\n    PackageLabel = \"{label}\"\r\nEnd Function\r\n")),
            new("modLicenseApiV1", VbaComponentKind.Standard, license),
            new("ThisWorkbook", VbaComponentKind.Workbook, WriteUtf8("ThisWorkbook.txt",
                "Option Explicit\r\nPrivate mReady As Boolean\r\n\r\nPrivate Sub Workbook_Open()\r\n    mReady = (Len(modReport.PackageLabel) > 0)\r\nEnd Sub\r\n" + events.Trim() + "\r\n")),
            // CodeName from an English template; the tab name lets it match "Trang_tính1" in Vietnamese Excel.
            new("Sheet1", VbaComponentKind.Sheet, WriteUtf8("Sheet1.txt",
                "Private Sub Worksheet_Change(ByVal Target As Range)\r\nEnd Sub\r\n"), SheetName: "Synthetic")
        };
        if (includeClass)
            sources.Add(new("clsRow", VbaComponentKind.Class, WriteAnsi("clsRow.cls",
                "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\nAttribute VB_Name = \"clsRow\"\r\nAttribute VB_GlobalNameSpace = False\r\nAttribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = False\r\nAttribute VB_Exposed = False\r\nOption Explicit\r\nPublic Value As Long\r\n")));

        var zip = Path.Combine(root, $"package-{label}.zip");
        var build = VbaPackageBuilder.Build(sources, zip);
        if (!build.IsValid) throw new InvalidOperationException(build.Error);
        var read = new VbaPackageReader().Read(zip, version, Path.Combine(root, "extract-" + label));
        return read.Package ?? throw new InvalidOperationException(read.Error);
    }

    /// <summary>Two sheets: "Synthetic" (matched by tab name) and "Dữ liệu" (an accented name for ChrW).</summary>
    private static string CreateWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = new Sheets();
        uint id = 1;
        foreach (var name in new[] { "Synthetic", "Dữ liệu" })
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = id++, Name = name });
        }
        workbookPart.Workbook.AppendChild(sheets);
        workbookPart.Workbook.Save();
        return path;
    }

    private static bool Check(string step, VbaPackageInstallResult result)
    {
        Expect(step, result.Succeeded, result.Message);
        return result.Succeeded;
    }

    private static void Expect(string step, bool ok, string detail)
    {
        if (!ok) _failures++;
        Console.WriteLine($"{(ok ? "PASS" : "FAIL")}: {step}{(ok ? "" : " — " + detail)}");
    }
}
