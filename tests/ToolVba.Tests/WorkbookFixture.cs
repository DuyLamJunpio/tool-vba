using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ToolVba.Tests;

internal static class WorkbookFixture
{
    public static readonly byte[] VbaBytes = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x01, 0x02, 0x03];

    public static string CreateDirectory() => Path.Combine(Path.GetTempPath(), $"tool-vba-test-{Guid.NewGuid():N}");

    public static void CreateWorkbook(string path, SpreadsheetDocumentType type, bool withVba = false)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var document = SpreadsheetDocument.Create(path, type);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        workbookPart.Workbook.AppendChild(new Sheets(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Synthetic"
        }));

        if (withVba)
        {
            var vbaPart = workbookPart.AddNewPart<VbaProjectPart>();
            using var stream = new MemoryStream(VbaBytes);
            vbaPart.FeedData(stream);
        }

        workbookPart.Workbook.Save();
    }

    public static string WriteVba(string path, byte[]? bytes = null)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, bytes ?? VbaBytes);
        return path;
    }

    public static string Hash(byte[] bytes) => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes));

    public static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream));
    }
}
