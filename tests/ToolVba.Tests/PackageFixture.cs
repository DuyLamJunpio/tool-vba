using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ToolVba.Core;

namespace ToolVba.Tests;

/// <summary>Synthetic VBA package helpers. No real company code is used in tests.</summary>
internal static class PackageFixture
{
    private static readonly Encoding Ansi = Encoding.Latin1;

    public static string StandardSource(string name, string body = "") =>
        $"Attribute VB_Name = \"{name}\"\r\nOption Explicit\r\n{body}";

    public static string ClassSource(string name) =>
        "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n" +
        $"Attribute VB_Name = \"{name}\"\r\nAttribute VB_GlobalNameSpace = False\r\nOption Explicit\r\n";

    public static string FormSource(string name) =>
        "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " + name + "\r\n" +
        $"   Caption = \"{name}\"\r\n   OleObjectBlob = \"{name}.frx\":0000\r\nEnd\r\n" +
        $"Attribute VB_Name = \"{name}\"\r\nOption Explicit\r\n";

    public const string WorkbookCode = "Option Explicit\r\n\r\nPrivate Sub Workbook_Open()\r\n    modReport.Prepare\r\nEnd Sub\r\n";

    /// <summary>Writes staged component files and builds a package with the production builder.</summary>
    public static string BuildPackage(TestDirectory temp, string zipName = "package.zip", bool includeForm = true, bool includeWorkbook = true)
    {
        var staging = Directory.CreateDirectory(temp.PathFor("staging-" + Guid.NewGuid().ToString("N"))).FullName;
        var sources = new List<VbaPackageSource>
        {
            new("modReport", VbaComponentKind.Standard, WriteAnsi(staging, "modReport.bas", StandardSource("modReport", "Public Sub Prepare()\r\nEnd Sub\r\n"))),
            new("clsRow", VbaComponentKind.Class, WriteAnsi(staging, "clsRow.cls", ClassSource("clsRow")))
        };
        if (includeForm)
        {
            var form = WriteAnsi(staging, "frmInput.frm", FormSource("frmInput"));
            var frx = Path.Combine(staging, "frmInput.frx");
            File.WriteAllBytes(frx, [1, 2, 3, 4]);
            sources.Add(new("frmInput", VbaComponentKind.Form, form, frx));
        }
        if (includeWorkbook)
        {
            var code = Path.Combine(staging, "ThisWorkbook.txt");
            File.WriteAllText(code, WorkbookCode, new UTF8Encoding(false));
            sources.Add(new("ThisWorkbook", VbaComponentKind.Workbook, code));
        }

        var zip = temp.PathFor(zipName);
        var result = VbaPackageBuilder.Build(sources, zip);
        Assert.True(result.IsValid, result.Error);
        return zip;
    }

    public static string WriteAnsi(string directory, string name, string content)
    {
        var path = Path.Combine(directory, name);
        File.WriteAllBytes(path, Ansi.GetBytes(content));
        return path;
    }

    public static string Sha256(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

    public static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexStringLower(SHA256.HashData(stream));
    }

    /// <summary>Writes a raw zip so tests can create packages the builder would refuse.</summary>
    public static string WriteRawZip(string path, IEnumerable<(string Name, byte[] Bytes)> entries)
    {
        using var stream = File.Create(path);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);
        foreach (var (name, bytes) in entries)
        {
            var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
            using var entryStream = entry.Open();
            entryStream.Write(bytes);
        }
        return path;
    }

    public static (string Name, byte[] Bytes) ManifestEntry(object manifest) =>
        ("package.json", JsonSerializer.SerializeToUtf8Bytes(manifest));

    public static (string Name, byte[] Bytes) ModuleEntry(string fileName, string moduleName, string body = "") =>
        (fileName, Ansi.GetBytes(StandardSource(moduleName, body)));

    public static object SingleModuleManifest(string fileName, byte[] bytes, string name = "modReport", string kind = "standard", int formatVersion = 1) => new
    {
        formatVersion,
        components = new[] { new { name, kind, file = fileName, sha256 = Sha256(bytes) } }
    };
}
