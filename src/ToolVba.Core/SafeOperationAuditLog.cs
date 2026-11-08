using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace ToolVba.Core;

public sealed class SafeOperationAuditLog
{
    private readonly string _logPath;
    private readonly BasModuleManifestValidator _moduleValidator;

    public SafeOperationAuditLog(string? logPath = null, BasModuleManifestValidator? moduleValidator = null)
    {
        _logPath = logPath ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ToolVba", "logs", "operations.jsonl");
        _moduleValidator = moduleValidator ?? new BasModuleManifestValidator();
    }

    public bool TryWrite(ExcelModuleInsertionRequest request, ExcelModuleInsertionResult result)
    {
        try
        {
            var module = _moduleValidator.Validate(request.ModulePath, request.ManifestPath);
            var record = new AuditRecord(
                DateTimeOffset.UtcNow,
                SafeFileName(request.InputWorkbookPath),
                TryHash(request.InputWorkbookPath),
                module.ModuleName,
                SafeFileName(request.ModulePath),
                TryHash(request.ModulePath),
                SafeFileName(request.OutputPath),
                result.Succeeded ? "PASS" : "FAIL",
                result.Message,
                ToolVersion);
            return Append(JsonSerializer.Serialize(record));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            return false;
        }
    }

    /// <summary>Records file names, hashes and the package version only; never workbook or VBA content.</summary>
    public bool TryWritePackageInstall(VbaPackageInstallRequest request, VbaPackageInstallResult result)
    {
        var record = new PackageInstallRecord(
            DateTimeOffset.UtcNow,
            SafeFileName(request.InputWorkbookPath),
            TryHash(request.InputWorkbookPath),
            request.Package.Version,
            SafeFileName(request.OutputPath),
            result.Succeeded ? "PASS" : "FAIL",
            result.Message,
            ToolVersion);
        return Append(JsonSerializer.Serialize(record));
    }

    private static string ToolVersion => typeof(SafeOperationAuditLog).Assembly.GetName().Version?.ToString() ?? "unknown";

    private bool Append(string line)
    {
        try
        {
            var directory = Path.GetDirectoryName(_logPath);
            if (string.IsNullOrWhiteSpace(directory)) return false;
            Directory.CreateDirectory(directory);
            File.AppendAllText(_logPath, line + Environment.NewLine, new UTF8Encoding(false));
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
        {
            return false;
        }
    }

    private sealed record PackageInstallRecord(
        DateTimeOffset TimestampUtc,
        string SourceFileName,
        string? SourceSha256,
        string PackageVersion,
        string OutputFileName,
        string Status,
        string Reason,
        string ToolVersion);

    private static string SafeFileName(string path)
    {
        try { return Path.GetFileName(path); }
        catch (ArgumentException) { return "invalid-path"; }
    }

    private static string? TryHash(string path)
    {
        try
        {
            using var stream = File.OpenRead(path);
            return Convert.ToHexString(SHA256.HashData(stream));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return null;
        }
    }

    private sealed record AuditRecord(
        DateTimeOffset TimestampUtc,
        string SourceFileName,
        string? SourceSha256,
        string? ModuleName,
        string ModuleFileName,
        string? ModuleSha256,
        string OutputFileName,
        string Status,
        string Reason,
        string ToolVersion);
}
