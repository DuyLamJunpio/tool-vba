namespace ToolVba.Core;

public sealed record InjectionRequest(
    string InputWorkbookPath,
    string VbaProjectPath,
    string OutputPath,
    bool ConfirmVbaReplacement = false,
    bool ConfirmOutputOverwrite = false);

public sealed record OperationProgress(string Stage, int Percent);

public sealed record OperationResult(
    bool Succeeded,
    string Message,
    string? OutputPath = null,
    IReadOnlyList<string>? Warnings = null)
{
    public static OperationResult Success(string outputPath, IReadOnlyList<string> warnings) =>
        new(true, "Đã tạo và kiểm tra workbook macro-enabled thành công.", outputPath, warnings);

    public static OperationResult Failure(string message) => new(false, message);
}

public sealed record WorkbookInspection(
    bool IsValid,
    string? Error,
    bool HasVbaProject,
    bool HasSignatureArtifacts,
    string? DocumentType);

public sealed record PackageValidationResult(bool IsValid, IReadOnlyList<string> Errors)
{
    public static PackageValidationResult Valid() => new(true, []);
    public static PackageValidationResult Invalid(params string[] errors) => new(false, errors);
}
