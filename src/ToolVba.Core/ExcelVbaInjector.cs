using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ToolVba.Core;

public sealed class ExcelVbaInjector
{
    public WorkbookInspection InspectWorkbook(string path)
    {
        var extension = Path.GetExtension(path);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            return new(false, "Chỉ hỗ trợ workbook Excel .xlsx hoặc .xlsm.", false, false, null);
        }

        if (!File.Exists(path))
            return new(false, "Không tìm thấy file Excel nguồn.", false, false, null);

        try
        {
            // Detect a workbook held exclusively by Excel before reporting a format error.
            using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read)) { }
            using var document = SpreadsheetDocument.Open(path, false);
            var workbookPart = document.WorkbookPart;
            if (workbookPart is null || document.DocumentType is not SpreadsheetDocumentType.Workbook and not SpreadsheetDocumentType.MacroEnabledWorkbook)
                return new(false, "File không phải workbook Excel Open XML được hỗ trợ.", false, false, document.DocumentType.ToString());

            var isXlsx = extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase);
            if (isXlsx && document.DocumentType != SpreadsheetDocumentType.Workbook ||
                !isXlsx && document.DocumentType != SpreadsheetDocumentType.MacroEnabledWorkbook)
            {
                return new(false, "Phần mở rộng không khớp với loại workbook bên trong file.", false, false, document.DocumentType.ToString());
            }

            var vbaCount = workbookPart.GetPartsOfType<VbaProjectPart>().Count();
            if (vbaCount > 1)
                return new(false, "Workbook chứa nhiều hơn một VBA project part và không thể xử lý an toàn.", false, false, document.DocumentType.ToString());
            var hasVba = vbaCount == 1;
            if (isXlsx && hasVba)
                return new(false, "File .xlsx có VBA project nhưng không được khai báo là macro-enabled. Hãy kiểm tra file nguồn trong Excel.", true, false, document.DocumentType.ToString());
            if (HasLegacyMacroSheets(path))
                return new(false, "Workbook có Excel 4.0 macro sheet. Tool dừng để tránh xử lý macro sheet ngoài phạm vi hỗ trợ.", hasVba, false, document.DocumentType.ToString());
            return new(true, null, hasVba, HasSignatureArtifacts(path), document.DocumentType.ToString());
        }
        catch (IOException ex) when ((ex.HResult & 0xFFFF) is 32 or 33)
        {
            return new(false, "Workbook nguồn đang được Excel hoặc ứng dụng khác sử dụng. Hãy đóng file nguồn rồi thử lại.", false, false, null);
        }
        catch (Exception ex) when (IsExpectedFileException(ex))
        {
            return new(false, "Không thể đọc file như workbook Excel Open XML. File có thể bị hỏng hoặc không đúng định dạng.", false, false, null);
        }
    }

    public OperationResult Inject(string stagedWorkbookPath, string vbaProjectPath, bool confirmVbaReplacement)
    {
        if (!File.Exists(vbaProjectPath))
            return OperationResult.Failure("Không tìm thấy file VBA project đã chọn.");
        if (!Path.GetExtension(vbaProjectPath).Equals(".bin", StringComparison.OrdinalIgnoreCase))
            return OperationResult.Failure("VBA project phải là file .bin.");

        try
        {
            using var vbaStream = new FileStream(vbaProjectPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (vbaStream.Length == 0)
                return OperationResult.Failure("File VBA project đang trống.");

            using var document = SpreadsheetDocument.Open(stagedWorkbookPath, true);
            var workbookPart = document.WorkbookPart;
            if (workbookPart is null)
                return OperationResult.Failure("Workbook không có phần dữ liệu Excel hợp lệ.");

            var existing = workbookPart.GetPartsOfType<VbaProjectPart>().SingleOrDefault();
            if (existing is not null && !confirmVbaReplacement)
                return OperationResult.Failure("Workbook đã có VBA project. Chọn “Thay VBA project hiện có” và xác nhận cảnh báo để tiếp tục.");

            // Reuse the existing part on replacement so its package URI remains xl/vbaProject.bin.
            var addedPart = existing ?? workbookPart.AddNewPart<VbaProjectPart>();
            vbaStream.Position = 0;
            addedPart.FeedData(vbaStream);
            document.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);
            return OperationResult.Success(stagedWorkbookPath, []);
        }
        catch (Exception ex) when (IsExpectedFileException(ex))
        {
            return OperationResult.Failure("Không thể chèn VBA vào bản tạm. Kiểm tra quyền truy cập và tính toàn vẹn của file.");
        }
    }

    internal static bool HasSignatureArtifacts(string path)
    {
        try
        {
            using var archive = ZipFile.OpenRead(path);
            return archive.Entries.Any(entry =>
                entry.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase) ||
                entry.FullName.StartsWith("xl/vbaProjectSignature", StringComparison.OrdinalIgnoreCase));
        }
        catch (Exception ex) when (IsExpectedFileException(ex))
        {
            return false;
        }
    }

    internal static bool HasLegacyMacroSheets(string path)
    {
        try
        {
            using var archive = ZipFile.OpenRead(path);
            return archive.Entries.Any(entry =>
                entry.FullName.StartsWith("xl/macrosheets/", StringComparison.OrdinalIgnoreCase) ||
                entry.FullName.StartsWith("xl/intlMacrosheets/", StringComparison.OrdinalIgnoreCase));
        }
        catch (Exception ex) when (IsExpectedFileException(ex))
        {
            return false;
        }
    }

    internal static bool IsExpectedFileException(Exception ex) =>
        ex is IOException or UnauthorizedAccessException or OpenXmlPackageException or ArgumentException or InvalidOperationException or
            InvalidDataException or NotSupportedException or System.IO.FileFormatException;
}
