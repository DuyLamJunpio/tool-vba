using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ToolVba.Core;

/// <summary>
/// End-to-end check against real Excel with synthetic files only:
/// install v1 into an .xlsx, add an employee module, update to v2,
/// and confirm a conflicting Workbook_Open is refused. Needs Excel and VBA project access.
/// </summary>
internal static class PackageSelfTest
{
    private static readonly Encoding Ansi = Encoding.Latin1;
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
        var host = new ExcelAutomationHost();
        IVbaWorkbookHost workbooks = host;
        var installer = new VbaPackageInstallService(workbooks);

        var source = CreateWorkbook(Path.Combine(root, "source.xlsx"));
        var v1 = BuildPackage(root, "v1", "2026.1.1.1", includeClass: true);
        var out1 = Path.Combine(root, "out1.xlsm");
        if (!Check("Cài v1 vào .xlsx", await installer.InstallAsync(new(source, out1, v1)))) return Finish(root);

        AddEmployeeModule(workbooks, out1, root);
        var v2 = BuildPackage(root, "v2", "2026.1.1.2", includeClass: false);
        var out2 = Path.Combine(root, "out2.xlsm");
        if (Check("Cập nhật lên v2 trên file đã có VBA", await installer.InstallAsync(new(out1, out2, v2))))
            VerifyUpdate(workbooks, out2);

        var conflicted = Path.Combine(root, "conflict.xlsm");
        File.Copy(out1, conflicted);
        using (var session = workbooks.Open(conflicted, readOnly: false))
        {
            session.SetDocumentCode(WorkbookModule(session), "Private Sub Workbook_Open()\r\n    MsgBox \"employee\"\r\nEnd Sub\r\n");
            session.Save();
        }
        var conflict = await installer.InstallAsync(new(conflicted, Path.Combine(root, "out3.xlsm"), v1));
        Expect("Từ chối Workbook_Open trùng", !conflict.Succeeded && conflict.Message.Contains("Workbook_Open"), conflict.Message);

        return Finish(root);
    }

    private static int Finish(string root)
    {
        Console.WriteLine(_failures == 0 ? "PASS: Self-test hoàn tất." : $"FAIL: {_failures} bước không đạt.");
        Console.WriteLine("Thư mục: " + root);
        return _failures == 0 ? 0 : 1;
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

    private static void AddEmployeeModule(IVbaWorkbookHost workbooks, string path, string root)
    {
        var file = Path.Combine(root, "modUser.bas");
        File.WriteAllBytes(file, Ansi.GetBytes("Attribute VB_Name = \"modUser\"\r\nOption Explicit\r\nPublic Sub Mine()\r\nEnd Sub\r\n"));
        using var session = workbooks.Open(path, readOnly: false);
        session.ImportComponent(file);
        session.Save();
    }

    private static string WorkbookModule(IVbaWorkbookSession session) =>
        session.GetComponents().First(c => c.Kind == VbaComponentKind.Workbook).Name;

    private static VbaPackage BuildPackage(string root, string label, string version, bool includeClass)
    {
        var staging = Directory.CreateDirectory(Path.Combine(root, "staging-" + label)).FullName;
        string WriteAnsi(string name, string text) { var p = Path.Combine(staging, name); File.WriteAllBytes(p, Ansi.GetBytes(text)); return p; }
        string WriteUtf8(string name, string text) { var p = Path.Combine(staging, name); File.WriteAllText(p, text, new UTF8Encoding(false)); return p; }

        var sources = new List<VbaPackageSource>
        {
            new("modReport", VbaComponentKind.Standard, WriteAnsi("modReport.bas",
                $"Attribute VB_Name = \"modReport\"\r\nOption Explicit\r\nPublic Function PackageLabel() As String\r\n    PackageLabel = \"{label}\"\r\nEnd Function\r\n")),
            new("ThisWorkbook", VbaComponentKind.Workbook, WriteUtf8("ThisWorkbook.txt",
                "Option Explicit\r\nPrivate mReady As Boolean\r\n\r\nPrivate Sub Workbook_Open()\r\n    mReady = (Len(modReport.PackageLabel) > 0)\r\nEnd Sub\r\n")),
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

    private static string CreateWorkbook(string path)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        workbookPart.Workbook.AppendChild(new Sheets(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Synthetic" }));
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
