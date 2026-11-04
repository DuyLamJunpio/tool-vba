using System.Security.Cryptography;
using System.Text.Json;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class SafeOperationAuditLogTests
{
    [Fact]
    public void Writes_only_file_names_hashes_status_and_safe_reason()
    {
        using var temp = new TestDirectory();
        var source = temp.PathFor("synthetic-source.xlsm");
        var module = temp.PathFor("approved-module.bas");
        var manifest = temp.PathFor("approved-module.json");
        var output = temp.PathFor("result.xlsm");
        var log = temp.PathFor("logs/operations.jsonl");
        const string sourceContent = "SYNTHETIC_WORKBOOK_CONTENT_MARKER";
        const string moduleContent = "SYNTHETIC_VBA_SOURCE_MARKER";
        File.WriteAllText(source, sourceContent);
        File.WriteAllText(module, moduleContent);
        var moduleHash = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(moduleContent)));
        File.WriteAllText(manifest, JsonSerializer.Serialize(new BasModuleManifest("SafeModule", moduleHash)));

        var wrote = new SafeOperationAuditLog(log).TryWrite(
            new ExcelModuleInsertionRequest(source, module, manifest, output),
            ExcelModuleInsertionResult.Failure("Test-only safe failure reason."));

        Assert.True(wrote);
        var line = File.ReadAllText(log);
        Assert.Contains("synthetic-source.xlsm", line);
        Assert.Contains("approved-module.bas", line);
        Assert.Contains(moduleHash, line);
        Assert.Contains("FAIL", line);
        Assert.DoesNotContain(sourceContent, line);
        Assert.DoesNotContain(moduleContent, line);
    }
}
