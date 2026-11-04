namespace ToolVba.Core;

public sealed record ExcelModuleInsertionRequest(
    string InputWorkbookPath,
    string ModulePath,
    string ManifestPath,
    string OutputPath,
    bool ConfirmOutputOverwrite = false);

public sealed record ExcelModuleInsertionResult(
    bool Succeeded,
    string Message,
    string? OutputPath = null,
    IReadOnlyList<string>? Warnings = null)
{
    public static ExcelModuleInsertionResult Success(string output, IReadOnlyList<string> warnings) =>
        new(true, "Đã thêm module và xác minh output thành công.", output, warnings);

    public static ExcelModuleInsertionResult Failure(string message) => new(false, message);
}

public interface IExcelAutomationHost
{
    IExcelAutomationSession Open(string workbookPath, bool readOnly);
}

public interface IExcelAutomationSession : IDisposable
{
    IReadOnlyList<string> GetComponentNames();
    string ImportStandardModule(string modulePath);
    void Save();
    void SaveAsMacroEnabled(string outputPath);
}

public sealed class ExcelAutomationException(string safeMessage, Exception? innerException = null)
    : Exception(safeMessage, innerException);

public sealed record ExcelModuleOperationProgress(string Stage, int Percent);
