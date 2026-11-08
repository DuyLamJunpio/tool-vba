using System.Reflection;
using System.Runtime.InteropServices;

namespace ToolVba.Core;

/// <summary>VBA project operations used by the package installer and exporter (late-bound Excel COM).</summary>
internal sealed partial class ExcelAutomationSession
{
    private const int StandardModuleType = 1;     // vbext_ct_StdModule
    private const int ClassModuleType = 2;        // vbext_ct_ClassModule
    private const int FormType = 3;               // vbext_ct_MSForm
    private const int DocumentType = 100;         // vbext_ct_Document
    private const int StringPropertyType = 4;     // msoPropertyTypeString
    private const int MaxComponentName = 31;

    public IReadOnlyList<VbaComponentInfo> GetComponents() => WithComponents(collection =>
    {
        var sheetNames = SheetNamesByCodeName();
        var raw = new List<(string Name, int Type)>();
        var count = (int)collection.Count;
        for (var index = 1; index <= count; index++)
        {
            object? component = null;
            try
            {
                component = collection.Item(index);
                dynamic item = component;
                raw.Add((ComponentName(item), (int)item.Type));
            }
            finally
            {
                ExcelAutomationHost.Release(component);
            }
        }

        // The workbook module is the document module that is not a sheet's CodeName. Workbook.CodeName
        // is only a tie-breaker: Excel keeps returning the old value after the module is renamed.
        var candidates = raw.Where(r => r.Type == DocumentType && !sheetNames.ContainsKey(r.Name)).Select(r => r.Name).ToList();
        var workbookModule = candidates.Count <= 1
            ? candidates.FirstOrDefault()
            : candidates.FirstOrDefault(c => string.Equals(c, WorkbookCodeName(), StringComparison.OrdinalIgnoreCase)) ?? candidates[0];
        return raw.Select(r =>
        {
            var kind = KindOf(r.Type, r.Name == workbookModule);
            return new VbaComponentInfo(r.Name, kind, kind == VbaComponentKind.Sheet && sheetNames.TryGetValue(r.Name, out var tab) ? tab : null);
        }).ToList();
    });

    public string ImportComponent(string filePath) => WithComponents(collection =>
    {
        object? imported = null;
        try
        {
            imported = collection.Import(filePath);
            return (string)((dynamic)imported).Name;
        }
        finally
        {
            ExcelAutomationHost.Release(imported);
        }
    });

    public void RemoveComponent(string name) => WithComponents(collection =>
    {
        object? component = null;
        try
        {
            component = FindComponent(collection, name);
            dynamic item = component;
            if ((int)item.Type == DocumentType)
                throw new ExcelAutomationException($"Không thể gỡ module {name} vì đó là ThisWorkbook hoặc sheet.");
            // Rename first: Excel may delay the removal, and a new import with the same name
            // would otherwise be renamed (for example to "Module11").
            item.Name = ("zzToolVbaOld" + Guid.NewGuid().ToString("N"))[..MaxComponentName];
            collection.Remove(component);
            return true;
        }
        finally
        {
            ExcelAutomationHost.Release(component);
        }
    });

    public void RenameDocumentModule(string name, string newName) => WithComponents(collection =>
    {
        object? component = null;
        object? properties = null;
        object? codeNameProperty = null;
        try
        {
            component = FindComponent(collection, name);
            // Setting VBComponent.Name would leave Workbook.CodeName stale, and Excel then writes the old
            // name to workbook.xml and recreates a broken module on the next open. "_CodeName" updates both.
            properties = ((dynamic)component).Properties;
            codeNameProperty = ((dynamic)properties).Item("_CodeName");
            ((dynamic)codeNameProperty).Value = newName;
            return true;
        }
        finally
        {
            ExcelAutomationHost.Release(codeNameProperty);
            ExcelAutomationHost.Release(properties);
            ExcelAutomationHost.Release(component);
        }
    });

    public string GetDocumentCode(string name) => WithCodeModule(name, codeModule =>
    {
        var count = (int)codeModule.CountOfLines;
        return count == 0 ? "" : (string)codeModule.Lines(1, count);
    });

    public void SetDocumentCode(string name, string code) => WithCodeModule(name, codeModule =>
    {
        var count = (int)codeModule.CountOfLines;
        if (count > 0) codeModule.DeleteLines(1, count);
        if (!string.IsNullOrEmpty(code)) codeModule.AddFromString(code);
        return true;
    });

    public void ExportComponent(string name, string filePath) => WithComponents(collection =>
    {
        object? component = null;
        try
        {
            component = FindComponent(collection, name);
            ((dynamic)component).Export(filePath);
            return true;
        }
        finally
        {
            ExcelAutomationHost.Release(component);
        }
    });

    public string? GetCustomProperty(string name) => WithCustomProperty(name, (_, property) =>
        property is null ? null : Convert.ToString(Invoke(property, "Value", BindingFlags.GetProperty), System.Globalization.CultureInfo.InvariantCulture));

    public void SetCustomProperty(string name, string value) => WithCustomProperty<object?>(name, (properties, property) =>
    {
        if (property is not null) Invoke(property, "Value", BindingFlags.SetProperty, value);
        else ExcelAutomationHost.Release(Invoke(properties, "Add", BindingFlags.InvokeMethod, name, false, StringPropertyType, value));
        return null;
    });

    public void DeleteCustomProperty(string name) => WithCustomProperty<object?>(name, (_, property) =>
    {
        if (property is not null) Invoke(property, "Delete", BindingFlags.InvokeMethod);
        return null;
    });

    private static VbaComponentKind KindOf(int type, bool isWorkbookModule) => type switch
    {
        StandardModuleType => VbaComponentKind.Standard,
        ClassModuleType => VbaComponentKind.Class,
        FormType => VbaComponentKind.Form,
        DocumentType => isWorkbookModule ? VbaComponentKind.Workbook : VbaComponentKind.Sheet,
        _ => VbaComponentKind.Other
    };

    /// <summary>The workbook's code name; read after the VBA project was touched so Excel has filled it in.</summary>
    private string WorkbookCodeName()
    {
        var codeName = (string?)((dynamic)RequireWorkbook()).CodeName;
        return string.IsNullOrEmpty(codeName) ? "ThisWorkbook" : codeName;
    }

    /// <summary>
    /// Excel can lose the Name of the workbook module, notably the localized default "SổlàmviệcNày" after
    /// an automated save as .xlsm; reading that module's CodeModule then crashes Excel. Such a module is
    /// reported with an empty name so callers can refuse to touch it.
    /// </summary>
    private static string ComponentName(dynamic item) => (string?)item.Name ?? "";

    /// <summary>Finds a component by walking the collection, so names Excel cannot look up still resolve.</summary>
    private object FindComponent(dynamic collection, string name)
    {
        var count = (int)collection.Count;
        for (var index = 1; index <= count; index++)
        {
            object component = collection.Item(index);
            if (string.Equals(ComponentName((dynamic)component), name, StringComparison.OrdinalIgnoreCase))
                return component;
            ExcelAutomationHost.Release(component);
        }
        throw new ExcelAutomationException($"Không tìm thấy component {name} trong VBA Project. File nguồn không bị sửa.");
    }

    private object RequireWorkbook() => _workbook ?? throw new ExcelAutomationException("Phiên Excel đã đóng.");

    /// <summary>Maps each sheet's CodeName (the document module name) to its tab name.</summary>
    private Dictionary<string, string> SheetNamesByCodeName()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        object? sheets = null;
        try
        {
            sheets = ((dynamic)RequireWorkbook()).Sheets;
            dynamic collection = sheets!;
            var count = (int)collection.Count;
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                try
                {
                    sheet = collection.Item(index);
                    dynamic item = sheet;
                    var codeName = (string?)item.CodeName;
                    if (!string.IsNullOrEmpty(codeName)) map[codeName] = (string)item.Name;
                }
                finally
                {
                    ExcelAutomationHost.Release(sheet);
                }
            }
            return map;
        }
        finally
        {
            ExcelAutomationHost.Release(sheets);
        }
    }

    /// <summary>
    /// Runs an action on VBComponents. Only a failure to reach the VBA project means access is not trusted;
    /// a failure inside the action is reported with the step and COM code so it is not mistaken for that.
    /// </summary>
    private T WithComponents<T>(Func<dynamic, T> action, [System.Runtime.CompilerServices.CallerMemberName] string caller = "")
    {
        object? project = null;
        object? components = null;
        try
        {
            try
            {
                project = ((dynamic)RequireWorkbook()).VBProject;
                components = ((dynamic)project).VBComponents;
            }
            catch (COMException ex)
            {
                throw VbaProjectAccessError(ex);
            }

            try
            {
                return action(components!);
            }
            catch (COMException ex)
            {
                throw new ExcelAutomationException(
                    $"Excel không {StepName(caller)} được (COM 0x{unchecked((uint)ex.HResult):X8}). Hãy đóng hết Excel rồi thử lại.", ex);
            }
        }
        finally
        {
            ExcelAutomationHost.Release(components);
            ExcelAutomationHost.Release(project);
        }
    }

    private static string StepName(string caller) => caller switch
    {
        nameof(GetComponents) => "đọc được danh sách module",
        nameof(ImportComponent) => "import module",
        nameof(RemoveComponent) => "gỡ module cũ",
        nameof(RenameDocumentModule) => "đổi tên module ThisWorkbook",
        nameof(ExportComponent) => "xuất module",
        _ => "đọc hoặc ghi code của ThisWorkbook/sheet"
    };

    private T WithCodeModule<T>(string name, Func<dynamic, T> action) => WithComponents(collection =>
    {
        object? component = null;
        object? codeModule = null;
        try
        {
            component = FindComponent(collection, name);
            codeModule = ((dynamic)component).CodeModule;
            return action(codeModule!);
        }
        finally
        {
            ExcelAutomationHost.Release(codeModule);
            ExcelAutomationHost.Release(component);
        }
    });

    /// <summary>
    /// DocumentProperties objects do not always bind through <c>dynamic</c>, so they are reached
    /// with IDispatch reflection calls instead.
    /// </summary>
    private T WithCustomProperty<T>(string name, Func<object, object?, T> action)
    {
        object? properties = null;
        object? match = null;
        try
        {
            properties = Invoke(RequireWorkbook(), "CustomDocumentProperties", BindingFlags.GetProperty)
                         ?? throw new ExcelAutomationException("Excel không đọc được thuộc tính tài liệu.");
            var count = Convert.ToInt32(Invoke(properties, "Count", BindingFlags.GetProperty), System.Globalization.CultureInfo.InvariantCulture);
            for (var index = 1; index <= count && match is null; index++)
            {
                var property = Invoke(properties, "Item", BindingFlags.GetProperty, index);
                if (property is not null && string.Equals((string?)Invoke(property, "Name", BindingFlags.GetProperty), name, StringComparison.OrdinalIgnoreCase))
                    match = property;
                else
                    ExcelAutomationHost.Release(property);
            }
            return action(properties, match);
        }
        catch (Exception ex) when (ex is COMException or TargetInvocationException)
        {
            throw new ExcelAutomationException("Excel không ghi được thông tin phiên bản code VBA vào file. File nguồn không bị sửa.", ex);
        }
        finally
        {
            ExcelAutomationHost.Release(match);
            ExcelAutomationHost.Release(properties);
        }
    }

    private static object? Invoke(object target, string member, BindingFlags flags, params object?[] args) =>
        target.GetType().InvokeMember(member, flags, null, target, args, System.Globalization.CultureInfo.InvariantCulture);
}
