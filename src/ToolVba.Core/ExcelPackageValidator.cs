using System.IO.Compression;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ToolVba.Core;

public sealed class ExcelPackageValidator
{
    private const string MacroWorkbookContentType = "application/vnd.ms-excel.sheet.macroEnabled.main";
    private const string VbaContentType = "application/vnd.ms-office.vbaProject";
    private const string VbaRelationshipType = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

    public PackageValidationResult Validate(string outputPath, string inputVbaPath)
    {
        var errors = new List<string>();
        if (!Path.GetExtension(outputPath).Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            errors.Add("File kết quả phải có phần mở rộng .xlsm.");
        try
        {
            using var document = SpreadsheetDocument.Open(outputPath, false);
            if (document.DocumentType != SpreadsheetDocumentType.MacroEnabledWorkbook)
                errors.Add("Document type không phải Macro-Enabled Workbook.");

            var workbookPart = document.WorkbookPart;
            if (workbookPart is null)
            {
                errors.Add("Không tìm thấy workbook part.");
                return new(false, errors);
            }

            if (!workbookPart.ContentType.StartsWith(MacroWorkbookContentType, StringComparison.OrdinalIgnoreCase))
                errors.Add("Content type của workbook không phải macro-enabled Excel.");

            var vbaParts = workbookPart.GetPartsOfType<VbaProjectPart>().ToArray();
            if (vbaParts.Length != 1)
                errors.Add("Workbook phải có đúng một VBA project part.");

            foreach (var part in vbaParts)
            {
                if (!string.Equals(part.ContentType, VbaContentType, StringComparison.OrdinalIgnoreCase))
                    errors.Add("Content type của VBA project không hợp lệ.");

                if (!workbookPart.Parts.Any(reference =>
                        ReferenceEquals(reference.OpenXmlPart, part) &&
                        string.Equals(reference.OpenXmlPart.RelationshipType, VbaRelationshipType, StringComparison.OrdinalIgnoreCase)))
                    errors.Add("Workbook chưa có relationship hợp lệ tới VBA project.");
            }

            using var archive = ZipFile.OpenRead(outputPath);
            if (archive.GetEntry("xl/vbaProject.bin") is null)
                errors.Add("Package thiếu entry xl/vbaProject.bin.");

            if (vbaParts.Length == 1 && File.Exists(inputVbaPath))
            {
                using var input = File.OpenRead(inputVbaPath);
                using var embedded = vbaParts[0].GetStream();
                var expectedHash = SHA256.HashData(input);
                var actualHash = SHA256.HashData(embedded);
                if (!CryptographicOperations.FixedTimeEquals(expectedHash, actualHash))
                    errors.Add("SHA-256 của VBA project trong output không khớp file đầu vào.");
            }
            else
            {
                errors.Add("Không thể so sánh VBA project với file đầu vào.");
            }
        }
        catch (Exception ex) when (ExcelVbaInjector.IsExpectedFileException(ex))
        {
            errors.Add("Không thể mở hoặc kiểm tra output như package Open XML hợp lệ.");
        }

        return errors.Count == 0 ? PackageValidationResult.Valid() : new(false, errors);
    }

    public PackageValidationResult ValidateMacroWorkbook(string workbookPath)
    {
        var errors = new List<string>();
        if (!Path.GetExtension(workbookPath).Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            errors.Add("File kết quả phải có phần mở rộng .xlsm.");

        try
        {
            using var document = SpreadsheetDocument.Open(workbookPath, false);
            if (document.DocumentType != SpreadsheetDocumentType.MacroEnabledWorkbook)
                errors.Add("Document type không phải Macro-Enabled Workbook.");

            var workbookPart = document.WorkbookPart;
            if (workbookPart is null)
            {
                errors.Add("Không tìm thấy workbook part.");
                return new(false, errors);
            }

            if (!workbookPart.ContentType.StartsWith(MacroWorkbookContentType, StringComparison.OrdinalIgnoreCase))
                errors.Add("Content type của workbook không phải macro-enabled Excel.");

            var vbaParts = workbookPart.GetPartsOfType<VbaProjectPart>().ToArray();
            if (vbaParts.Length != 1)
                errors.Add("Workbook phải có đúng một VBA project part.");
            foreach (var part in vbaParts)
            {
                if (!string.Equals(part.ContentType, VbaContentType, StringComparison.OrdinalIgnoreCase))
                    errors.Add("Content type của VBA project không hợp lệ.");
                if (!workbookPart.Parts.Any(reference =>
                        ReferenceEquals(reference.OpenXmlPart, part) &&
                        string.Equals(reference.OpenXmlPart.RelationshipType, VbaRelationshipType, StringComparison.OrdinalIgnoreCase)))
                    errors.Add("Workbook chưa có relationship hợp lệ tới VBA project.");
            }

            using var archive = ZipFile.OpenRead(workbookPath);
            if (archive.GetEntry("xl/vbaProject.bin") is null)
                errors.Add("Package thiếu entry xl/vbaProject.bin.");
        }
        catch (Exception ex) when (ExcelVbaInjector.IsExpectedFileException(ex))
        {
            errors.Add("Không thể mở hoặc kiểm tra output như package Open XML hợp lệ.");
        }

        return errors.Count == 0 ? PackageValidationResult.Valid() : new(false, errors);
    }
}
