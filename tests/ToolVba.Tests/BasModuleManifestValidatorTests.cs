using System.Security.Cryptography;
using System.Text.Json;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class BasModuleManifestValidatorTests
{
    [Fact]
    public void Accepts_module_when_sha256_matches_manifest()
    {
        using var temp = new TestDirectory();
        var module = WriteModule(temp.PathFor("Approved.bas"));
        var manifest = WriteManifest(temp.PathFor("Approved.manifest.json"), "Approved", Hash(module));

        var result = new BasModuleManifestValidator().Validate(module, manifest);

        Assert.True(result.IsValid, result.Error);
        Assert.Equal("Approved", result.ModuleName);
    }

    [Fact]
    public void Rejects_module_when_sha256_does_not_match_manifest()
    {
        using var temp = new TestDirectory();
        var module = WriteModule(temp.PathFor("Approved.bas"));
        var manifest = WriteManifest(temp.PathFor("Approved.manifest.json"), "Approved", new string('0', 64));

        var result = new BasModuleManifestValidator().Validate(module, manifest);

        Assert.False(result.IsValid);
        Assert.Contains("SHA-256", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rejects_malformed_manifest_without_exposing_exception_details()
    {
        using var temp = new TestDirectory();
        var module = WriteModule(temp.PathFor("Approved.bas"));
        var manifest = temp.PathFor("bad.manifest.json");
        File.WriteAllText(manifest, "{");

        var result = new BasModuleManifestValidator().Validate(module, manifest);

        Assert.False(result.IsValid);
        Assert.NotNull(result.Error);
        Assert.DoesNotContain("JsonException", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rejects_reserved_input_identifier_even_when_manifest_hash_matches()
    {
        using var temp = new TestDirectory();
        var module = temp.PathFor("Approved.bas");
        File.WriteAllText(module,
            "Attribute VB_Name = \"Approved\"\r\nOption Explicit\r\n" +
            "Private Declare PtrSafe Function NativeCall Lib \"example.dll\" ( _\r\n" +
            "    ByVal input As LongPtr) As Long\r\n",
            new System.Text.UTF8Encoding(false));
        var manifest = WriteManifest(temp.PathFor("Approved.manifest.json"), "Approved", Hash(module));

        var result = new BasModuleManifestValidator().Validate(module, manifest);

        Assert.False(result.IsValid);
        Assert.Contains("Input", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public void Shipped_license_module_matches_its_manifest()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "ToolVba.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        var sample = Path.Combine(directory!.FullName, "license-sample", "api-v1-one-file");

        var result = new BasModuleManifestValidator().Validate(
            Path.Combine(sample, "modLicenseApiV1.bas"),
            Path.Combine(sample, "modLicenseApiV1.manifest.json"));

        Assert.True(result.IsValid, result.Error);
        Assert.Equal("modLicenseApiV1", result.ModuleName);
    }

    private static string WriteModule(string path)
    {
        var source = "Attribute VB_Name = \"Approved\"\r\nOption Explicit\r\nPublic Sub Smoke()\r\n    MsgBox \"OK\"\r\nEnd Sub\r\n";
        File.WriteAllText(path, source, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        return path;
    }

    private static string WriteManifest(string path, string name, string hash)
    {
        File.WriteAllText(path, JsonSerializer.Serialize(new BasModuleManifest(name, hash)));
        return path;
    }

    private static string Hash(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
