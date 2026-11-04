using System.Runtime.InteropServices;

namespace ToolVba.Core;

public sealed class ExcelAutomationHost : IExcelAutomationHost
{
    public IExcelAutomationSession Open(string workbookPath, bool readOnly)
    {
        if (!OperatingSystem.IsWindows())
            throw new ExcelAutomationException("Chế độ thêm module .bas chỉ chạy trên Windows có Microsoft Excel desktop.");

        var excelType = Type.GetTypeFromProgID("Excel.Application", throwOnError: false);
        if (excelType is null)
            throw new ExcelAutomationException("Không tìm thấy Microsoft Excel desktop trên máy này.");

        object? application = null;
        object? workbooks = null;
        object? workbook = null;
        object? scratchWorkbook = null;
        var stage = "khởi động Excel";
        try
        {
            application = Activator.CreateInstance(excelType)
                ?? throw new ExcelAutomationException("Không thể khởi động Microsoft Excel.");
            dynamic excel = application;
            stage = "ẩn cửa sổ Excel";
            excel.Visible = false;
            stage = "tắt hộp thoại Excel";
            excel.DisplayAlerts = false;
            stage = "tắt sự kiện workbook";
            excel.EnableEvents = false;
            stage = "tắt hỏi cập nhật liên kết";
            excel.AskToUpdateLinks = false;
            stage = "vô hiệu macro cho file mở bằng Automation";
            excel.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
            stage = "tạo workbook trống trong phiên Excel";
            workbooks = excel.Workbooks;
            dynamic collection = workbooks;
            scratchWorkbook = collection.Add(-4167); // xlWBATWorksheet, no template file
            stage = "đặt chế độ tính toán thủ công";
            excel.Calculation = -4135; // xlCalculationManual
            stage = "tắt tính toán trước khi lưu";
            excel.CalculateBeforeSave = false;
            stage = "vô hiệu macro trước khi mở bản tạm";
            excel.AutomationSecurity = 3;

            stage = "mở bản tạm trong Excel";
            workbook = collection.Open(workbookPath, 0, readOnly);
            stage = "đóng workbook trống";
            ((dynamic)scratchWorkbook).Close(false);
            Release(scratchWorkbook);
            scratchWorkbook = null;
            return new ExcelAutomationSession(application, workbooks, workbook);
        }
        catch (ExcelAutomationException)
        {
            CloseAndRelease(application, workbooks, workbook, scratchWorkbook);
            throw;
        }
        catch (COMException ex)
        {
            CloseAndRelease(application, workbooks, workbook, scratchWorkbook);
            throw new ExcelAutomationException(DescribeComFailure(stage, ex.HResult), ex);
        }
        catch (Exception ex) when (ex is InvalidOperationException or UnauthorizedAccessException or ArgumentException)
        {
            CloseAndRelease(application, workbooks, workbook, scratchWorkbook);
            throw new ExcelAutomationException($"Không thể {stage}. Kiểm tra Excel và quyền truy cập workbook.", ex);
        }
    }

    private static string DescribeComFailure(string stage, int hresult)
    {
        var code = $"0x{unchecked((uint)hresult):X8}";
        var guidance = unchecked((uint)hresult) switch
        {
            0x80070520 => "Windows không có phiên đăng nhập hợp lệ cho Excel. Hãy lưu công việc, đóng công cụ và mở lại tool-vba.exe trực tiếp từ File Explorer trong phiên Windows đang đăng nhập.",
            0x80040154 => "Excel Automation chưa được đăng ký trên máy. Hãy kiểm tra Excel desktop có mở được từ menu Start.",
            0x80070005 => "Windows từ chối quyền khởi động hoặc điều khiển Excel. Hãy mở Excel trực tiếp và kiểm tra chính sách của máy.",
            0x80080005 => "Excel không hoàn tất khởi động qua Automation. Hãy đóng Excel, mở thử Excel trực tiếp rồi chạy lại công cụ.",
            _ => "Hãy mở thử Excel trực tiếp trong Windows và gửi mã lỗi này nếu cần hỗ trợ tiếp."
        };
        return $"Excel gặp lỗi khi {stage} (COM {code}). {guidance} File nguồn không bị sửa.";
    }

    private static void CloseAndRelease(object? application, object? workbooks, object? workbook, object? scratchWorkbook)
    {
        if (workbook is not null)
        {
            try { ((dynamic)workbook).Close(false); } catch { }
            Release(workbook);
        }
        if (scratchWorkbook is not null)
        {
            try { ((dynamic)scratchWorkbook).Close(false); } catch { }
            Release(scratchWorkbook);
        }
        if (application is not null)
        {
            try { ((dynamic)application).Quit(); } catch { }
        }
        Release(workbooks);
        Release(application);
    }

    internal static void Release(object? value)
    {
        if (OperatingSystem.IsWindows() && value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.FinalReleaseComObject(value); } catch (InvalidComObjectException) { }
            catch (COMException) { }
        }
    }
}

internal sealed class ExcelAutomationSession : IExcelAutomationSession
{
    private object? _application;
    private object? _workbooks;
    private object? _workbook;

    public ExcelAutomationSession(object application, object workbooks, object workbook)
    {
        _application = application;
        _workbooks = workbooks;
        _workbook = workbook;
    }

    public IReadOnlyList<string> GetComponentNames()
    {
        object? project = null;
        object? components = null;
        try
        {
            dynamic workbook = _workbook ?? throw new ExcelAutomationException("Phiên Excel đã đóng.");
            project = workbook.VBProject;
            dynamic vbaProject = project;
            components = vbaProject.VBComponents;
            dynamic collection = components;
            var names = new List<string>();
            var count = (int)collection.Count;
            for (var index = 1; index <= count; index++)
            {
                object? component = null;
                try
                {
                    component = collection.Item(index);
                    names.Add((string)((dynamic)component).Name);
                }
                finally
                {
                    ExcelAutomationHost.Release(component);
                }
            }
            return names;
        }
        catch (COMException ex)
        {
            throw VbaProjectAccessError(ex);
        }
        finally
        {
            ExcelAutomationHost.Release(components);
            ExcelAutomationHost.Release(project);
        }
    }

    public string ImportStandardModule(string modulePath)
    {
        object? project = null;
        object? components = null;
        object? imported = null;
        try
        {
            dynamic workbook = _workbook ?? throw new ExcelAutomationException("Phiên Excel đã đóng.");
            project = workbook.VBProject;
            dynamic vbaProject = project;
            components = vbaProject.VBComponents;
            dynamic collection = components;
            imported = collection.Import(modulePath);
            return (string)((dynamic)imported).Name;
        }
        catch (COMException ex)
        {
            throw VbaProjectAccessError(ex);
        }
        finally
        {
            ExcelAutomationHost.Release(imported);
            ExcelAutomationHost.Release(components);
            ExcelAutomationHost.Release(project);
        }
    }

    public void Save()
    {
        try
        {
            dynamic workbook = _workbook ?? throw new ExcelAutomationException("Phiên Excel đã đóng.");
            workbook.Save();
        }
        catch (COMException ex)
        {
            throw new ExcelAutomationException("Excel không lưu được workbook tạm. File nguồn không bị sửa.", ex);
        }
    }

    public void SaveAsMacroEnabled(string outputPath)
    {
        try
        {
            dynamic workbook = _workbook ?? throw new ExcelAutomationException("Phiên Excel đã đóng.");
            workbook.SaveAs(outputPath, 52); // xlOpenXMLWorkbookMacroEnabled
        }
        catch (COMException ex)
        {
            throw new ExcelAutomationException("Excel không lưu được workbook tạm dạng .xlsm. File nguồn không bị sửa.", ex);
        }
    }

    public void Dispose()
    {
        var workbook = _workbook;
        var application = _application;
        if (workbook is not null)
        {
            try { ((dynamic)workbook).Close(false); } catch { }
            ExcelAutomationHost.Release(workbook);
            _workbook = null;
        }
        if (application is not null)
        {
            try { ((dynamic)application).Quit(); } catch { }
            ExcelAutomationHost.Release(_workbooks);
            ExcelAutomationHost.Release(application);
            _workbooks = null;
            _application = null;
        }
    }

    private static ExcelAutomationException VbaProjectAccessError(COMException ex) =>
        new("Excel từ chối truy cập VBA Project. Quản trị viên cần bật “Trust access to the VBA project object model” trên máy đóng gói theo chính sách tổ chức; công cụ không tự đổi Trust Center hoặc Registry.", ex);
}
