namespace ToolVba.Core;

/// <param name="SheetName">Worksheet tab name for sheet document modules; null otherwise.</param>
public sealed record VbaComponentInfo(string Name, VbaComponentKind Kind, string? SheetName = null);

/// <summary>Opens workbooks in a hidden Excel instance with macros and events disabled.</summary>
public interface IVbaWorkbookHost
{
    IVbaWorkbookSession Open(string workbookPath, bool readOnly);
}

/// <summary>VBA project operations needed to install or export a company package.</summary>
public interface IVbaWorkbookSession : IDisposable
{
    IReadOnlyList<VbaComponentInfo> GetComponents();

    /// <summary>Imports a .bas, .cls or .frm file and returns the component name Excel assigned.</summary>
    string ImportComponent(string filePath);

    /// <summary>Removes a standard, class or form component. Document modules cannot be removed.</summary>
    void RemoveComponent(string name);

    /// <summary>Changes the CodeName of a document module (ThisWorkbook or a sheet) and keeps Excel's copy in sync.</summary>
    void RenameDocumentModule(string name, string newName);

    string GetDocumentCode(string name);
    void SetDocumentCode(string name, string code);
    void ExportComponent(string name, string filePath);
    string? GetCustomProperty(string name);
    void SetCustomProperty(string name, string value);
    void DeleteCustomProperty(string name);
    void Save();
    void SaveAsMacroEnabled(string outputPath);

    /// <summary>Every sheet of the workbook in tab order, worksheets and chart sheets.</summary>
    IReadOnlyList<WorkbookSheetInfo> GetSheets();

    /// <summary>True when Review &gt; Protect Workbook locks the structure (sheets cannot be added or hidden).</summary>
    bool IsStructureProtected();

    /// <summary>
    /// Puts <paramref name="lines"/> in column A of the worksheet <paramref name="name"/> (added first when
    /// <paramref name="create"/>), makes it the first, visible and active sheet.
    /// </summary>
    void WriteGuideSheet(string name, IReadOnlyList<string> lines, bool create);

    void SetSheetVisibility(string name, SheetVisibility visibility);

    /// <summary>The text of one cell (for example "A1") of a sheet; "" when empty.</summary>
    string ReadCell(string sheetName, string address);
}
