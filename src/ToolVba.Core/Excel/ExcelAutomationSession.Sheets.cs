using System.Runtime.InteropServices;

namespace ToolVba.Core;

/// <summary>Sheet operations for the licensing guide sheet and the protected sheets (late-bound Excel COM).</summary>
internal sealed partial class ExcelAutomationSession
{
    private const int XlSheetVisible = -1;
    private const int XlSheetHidden = 0;
    private const int XlSheetVeryHidden = 2;
    private const double GuideColumnWidth = 120;

    public IReadOnlyList<WorkbookSheetInfo> GetSheets() => WithSheets("đọc danh sách sheet", workbook =>
    {
        // Cast first: a dynamic argument would make the call dynamic, and extension methods do not bind on dynamic.
        List<string> worksheets = Names((object)workbook.Worksheets);
        var worksheetNames = worksheets.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var result = new List<WorkbookSheetInfo>();
        object? sheets = null;
        try
        {
            sheets = workbook.Sheets;
            dynamic collection = sheets!;
            var count = (int)collection.Count;
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                try
                {
                    sheet = collection.Item(index);
                    dynamic item = sheet;
                    var name = (string)item.Name;
                    result.Add(new(name, !worksheetNames.Contains(name), VisibilityOf((int)item.Visible)));
                }
                finally
                {
                    ExcelAutomationHost.Release(sheet);
                }
            }
            return result;
        }
        finally
        {
            ExcelAutomationHost.Release(sheets);
        }
    });

    public bool IsStructureProtected() => WithSheets("đọc trạng thái khóa cấu trúc", workbook => (bool)workbook.ProtectStructure);

    public void WriteGuideSheet(string name, IReadOnlyList<string> lines, bool create) => WithSheets("tạo sheet hướng dẫn", workbook =>
    {
        object? worksheets = null;
        object? first = null;
        object? sheet = null;
        try
        {
            worksheets = workbook.Worksheets;
            dynamic collection = worksheets!;
            first = collection.Item(1);
            sheet = create ? collection.Add(Before: first) : collection.Item(name);
            dynamic guide = sheet!;
            if (create) guide.Name = name;
            guide.Visible = XlSheetVisible;
            if (!create && (int)guide.Index != 1) guide.Move(Before: first);
            ClearCells(guide);
            for (var row = 0; row < lines.Count; row++)
                SetCell(guide, $"A{row + 1}", lines[row]);
            SetColumnWidth(guide);
            guide.Activate();
            return true;
        }
        finally
        {
            ExcelAutomationHost.Release(sheet);
            ExcelAutomationHost.Release(first);
            ExcelAutomationHost.Release(worksheets);
        }
    });

    public void SetSheetVisibility(string name, SheetVisibility visibility) => WithSheets("ẩn sheet", workbook =>
    {
        object? sheets = null;
        object? sheet = null;
        try
        {
            sheets = workbook.Sheets;
            sheet = ((dynamic)sheets!).Item(name);
            ((dynamic)sheet!).Visible = visibility switch
            {
                SheetVisibility.Visible => XlSheetVisible,
                SheetVisibility.Hidden => XlSheetHidden,
                _ => XlSheetVeryHidden
            };
            return true;
        }
        finally
        {
            ExcelAutomationHost.Release(sheet);
            ExcelAutomationHost.Release(sheets);
        }
    });

    public string ReadCell(string sheetName, string address) => WithSheets("đọc sheet hướng dẫn", workbook =>
    {
        object? sheets = null;
        object? sheet = null;
        object? range = null;
        try
        {
            sheets = workbook.Sheets;
            sheet = ((dynamic)sheets!).Item(sheetName);
            range = ((dynamic)sheet!).Range(address);
            return Convert.ToString(((dynamic)range!).Value2, System.Globalization.CultureInfo.InvariantCulture) ?? "";
        }
        finally
        {
            ExcelAutomationHost.Release(range);
            ExcelAutomationHost.Release(sheet);
            ExcelAutomationHost.Release(sheets);
        }
    });

    private static void ClearCells(dynamic sheet)
    {
        object? cells = null;
        try
        {
            cells = sheet.Cells;
            ((dynamic)cells!).Clear();
        }
        finally
        {
            ExcelAutomationHost.Release(cells);
        }
    }

    private static SheetVisibility VisibilityOf(int value) => value switch
    {
        XlSheetVisible => SheetVisibility.Visible,
        XlSheetVeryHidden => SheetVisibility.VeryHidden,
        _ => SheetVisibility.Hidden
    };

    private static void SetCell(dynamic sheet, string address, string value)
    {
        object? range = null;
        try
        {
            range = sheet.Range(address);
            ((dynamic)range!).Value2 = value;
        }
        finally
        {
            ExcelAutomationHost.Release(range);
        }
    }

    private static void SetColumnWidth(dynamic sheet)
    {
        object? column = null;
        try
        {
            column = sheet.Range("A:A");
            ((dynamic)column!).ColumnWidth = GuideColumnWidth;
        }
        finally
        {
            ExcelAutomationHost.Release(column);
        }
    }

    private static List<string> Names(object? collectionObject)
    {
        var names = new List<string>();
        try
        {
            dynamic collection = collectionObject!;
            var count = (int)collection.Count;
            for (var index = 1; index <= count; index++)
            {
                object? item = null;
                try
                {
                    item = collection.Item(index);
                    names.Add((string)((dynamic)item!).Name);
                }
                finally
                {
                    ExcelAutomationHost.Release(item);
                }
            }
            return names;
        }
        finally
        {
            ExcelAutomationHost.Release(collectionObject);
        }
    }

    private T WithSheets<T>(string step, Func<dynamic, T> action)
    {
        try
        {
            return action(RequireWorkbook());
        }
        catch (COMException ex)
        {
            throw new ExcelAutomationException(
                $"Excel không {step} được (COM 0x{unchecked((uint)ex.HResult):X8}). Hãy đóng hết Excel rồi thử lại. File nguồn không bị sửa.", ex);
        }
    }
}
