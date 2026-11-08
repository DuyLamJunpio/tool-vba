using System.Text;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class VbaPackageReaderTests
{
    [Fact]
    public void Reads_valid_package_with_all_component_kinds_and_extracts_verified_files()
    {
        using var temp = new TestDirectory();
        var zip = PackageFixture.BuildPackage(temp);
        var extract = temp.PathFor("extract");

        var result = new VbaPackageReader().Read(zip, "2026.11.7.1", extract);

        Assert.True(result.IsValid, result.Error);
        var package = result.Package!;
        Assert.Equal("2026.11.7.1", package.Version);
        Assert.Equal(
            ["modReport", "clsRow", "frmInput", "ThisWorkbook"],
            package.Components.Select(c => c.Name));
        var form = package.Components.Single(c => c.Kind == VbaComponentKind.Form);
        Assert.True(File.Exists(form.FilePath));
        Assert.True(File.Exists(form.ExtraFilePath));
        Assert.Equal(Path.GetDirectoryName(form.FilePath), Path.GetDirectoryName(form.ExtraFilePath));
        Assert.All(package.Components, c => Assert.StartsWith(Path.GetFullPath(extract), c.FilePath));
    }

    [Theory]
    [InlineData("../evil.bas")]
    [InlineData("sub/evil.bas")]
    [InlineData("C:evil.bas")]
    [InlineData("..")]
    public void Rejects_entries_that_are_not_flat_safe_file_names(string entryName)
    {
        using var temp = new TestDirectory();
        var (name, bytes) = PackageFixture.ModuleEntry("modReport.bas", "modReport");
        var zip = PackageFixture.WriteRawZip(temp.PathFor("bad.zip"),
        [
            PackageFixture.ManifestEntry(PackageFixture.SingleModuleManifest(name, bytes)),
            (name, bytes),
            (entryName, [1])
        ]);

        var result = new VbaPackageReader().Read(zip, "1.0.0.1", temp.PathFor("extract"));

        Assert.False(result.IsValid);
        Assert.Contains("tên file", result.Error, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(Path.Combine(temp.Root, "evil.bas")));
    }

    [Theory]
    [InlineData("NUL.bas")]
    [InlineData("con.txt")]
    [InlineData("COM1")]
    public void Rejects_windows_device_names(string entryName)
    {
        using var temp = new TestDirectory();
        var (name, bytes) = PackageFixture.ModuleEntry("modReport.bas", "modReport");
        var zip = PackageFixture.WriteRawZip(temp.PathFor("bad.zip"),
        [
            PackageFixture.ManifestEntry(PackageFixture.SingleModuleManifest(name, bytes)),
            (name, bytes),
            (entryName, [1])
        ]);

        Assert.False(new VbaPackageReader().Read(zip, "1.0.0.1", temp.PathFor("extract")).IsValid);
    }

    [Fact]
    public void Rejects_unknown_or_differently_cased_manifest_fields()
    {
        using var temp = new TestDirectory();
        var (name, bytes) = PackageFixture.ModuleEntry("modReport.bas", "modReport");
        var component = new { name = "modReport", kind = "standard", file = name, sha256 = PackageFixture.Sha256(bytes) };
        var unknown = PackageFixture.WriteRawZip(temp.PathFor("a.zip"),
            [PackageFixture.ManifestEntry(new { formatVersion = 1, components = new[] { component }, extra = true }), (name, bytes)]);
        var cased = PackageFixture.WriteRawZip(temp.PathFor("b.zip"),
            [PackageFixture.ManifestEntry(new { formatVersion = 1, Components = new[] { component } }), (name, bytes)]);

        Assert.False(new VbaPackageReader().Read(unknown, "1.0.0.1", temp.PathFor("x1")).IsValid);
        Assert.False(new VbaPackageReader().Read(cased, "1.0.0.1", temp.PathFor("x2")).IsValid);
    }

    [Fact]
    public void Rejects_hash_mismatch()
    {
        using var temp = new TestDirectory();
        var (name, bytes) = PackageFixture.ModuleEntry("modReport.bas", "modReport");
        var zip = PackageFixture.WriteRawZip(temp.PathFor("bad.zip"),
        [
            PackageFixture.ManifestEntry(PackageFixture.SingleModuleManifest(name, Encoding.ASCII.GetBytes("other"))),
            (name, bytes)
        ]);

        var result = new VbaPackageReader().Read(zip, "1.0.0.1", temp.PathFor("extract"));

        Assert.False(result.IsValid);
        Assert.Contains("SHA-256", result.Error);
    }

    [Fact]
    public void Rejects_missing_manifest_and_unsupported_format_version()
    {
        using var temp = new TestDirectory();
        var (name, bytes) = PackageFixture.ModuleEntry("modReport.bas", "modReport");
        var noManifest = PackageFixture.WriteRawZip(temp.PathFor("a.zip"), [(name, bytes)]);
        var wrongVersion = PackageFixture.WriteRawZip(temp.PathFor("b.zip"),
        [
            PackageFixture.ManifestEntry(PackageFixture.SingleModuleManifest(name, bytes, formatVersion: 2)),
            (name, bytes)
        ]);

        Assert.Contains("package.json", new VbaPackageReader().Read(noManifest, "1.0.0.1", temp.PathFor("x1")).Error);
        Assert.Contains("định dạng", new VbaPackageReader().Read(wrongVersion, "1.0.0.1", temp.PathFor("x2")).Error);
    }

    [Theory]
    [InlineData("modReport.cls", "standard")]
    [InlineData("modReport.bas", "class")]
    [InlineData("modReport.bas", "workbook")]
    public void Rejects_file_extension_that_does_not_match_component_kind(string fileName, string kind)
    {
        using var temp = new TestDirectory();
        var (_, bytes) = PackageFixture.ModuleEntry(fileName, "modReport");
        var zip = PackageFixture.WriteRawZip(temp.PathFor("bad.zip"),
        [
            PackageFixture.ManifestEntry(PackageFixture.SingleModuleManifest(fileName, bytes, kind: kind)),
            (fileName, bytes)
        ]);

        var result = new VbaPackageReader().Read(zip, "1.0.0.1", temp.PathFor("extract"));

        Assert.False(result.IsValid);
        Assert.Contains("phần mở rộng", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Rejects_form_without_frx_and_duplicate_names_ignoring_case()
    {
        using var temp = new TestDirectory();
        var form = Encoding.Latin1.GetBytes(PackageFixture.FormSource("frmInput"));
        var noFrx = PackageFixture.WriteRawZip(temp.PathFor("a.zip"),
        [
            PackageFixture.ManifestEntry(PackageFixture.SingleModuleManifest("frmInput.frm", form, "frmInput", "form")),
            ("frmInput.frm", form)
        ]);
        var (name, bytes) = PackageFixture.ModuleEntry("modReport.bas", "modReport");
        var duplicate = PackageFixture.WriteRawZip(temp.PathFor("b.zip"),
        [
            PackageFixture.ManifestEntry(new
            {
                formatVersion = 1,
                components = new[]
                {
                    new { name = "modReport", kind = "standard", file = name, sha256 = PackageFixture.Sha256(bytes) },
                    new { name = "MODREPORT", kind = "standard", file = name, sha256 = PackageFixture.Sha256(bytes) }
                }
            }),
            (name, bytes)
        ]);

        Assert.Contains(".frx", new VbaPackageReader().Read(noFrx, "1.0.0.1", temp.PathFor("x1")).Error);
        Assert.Contains("trùng", new VbaPackageReader().Read(duplicate, "1.0.0.1", temp.PathFor("x2")).Error);
    }

    [Fact]
    public void Accepts_localized_sheet_code_name_with_ascii_file_and_tab_name()
    {
        using var temp = new TestDirectory();
        var code = Encoding.UTF8.GetBytes("Private Sub Worksheet_Change(ByVal Target As Range)\r\nEnd Sub\r\n");
        var zip = PackageFixture.WriteRawZip(temp.PathFor("vi.zip"),
        [
            PackageFixture.ManifestEntry(new
            {
                formatVersion = 1,
                components = new[] { new { name = "Trang_tính1", kind = "sheet", file = "Sheet1.txt", sha256 = PackageFixture.Sha256(code), sheetName = "Báo cáo" } }
            }),
            ("Sheet1.txt", code)
        ]);

        var result = new VbaPackageReader().Read(zip, "1.0.0.1", temp.PathFor("extract"));

        Assert.True(result.IsValid, result.Error);
        Assert.Equal("Báo cáo", result.Package!.Components[0].SheetName);
    }

    [Fact]
    public void Rejects_sheet_name_on_non_sheet_component()
    {
        using var temp = new TestDirectory();
        var (name, bytes) = PackageFixture.ModuleEntry("modReport.bas", "modReport");
        var zip = PackageFixture.WriteRawZip(temp.PathFor("bad.zip"),
        [
            PackageFixture.ManifestEntry(new
            {
                formatVersion = 1,
                components = new[] { new { name = "modReport", kind = "standard", file = name, sha256 = PackageFixture.Sha256(bytes), sheetName = "x" } }
            }),
            (name, bytes)
        ]);

        Assert.False(new VbaPackageReader().Read(zip, "1.0.0.1", temp.PathFor("extract")).IsValid);
    }

    [Fact]
    public void Rejects_module_whose_vb_name_differs_from_manifest()
    {
        using var temp = new TestDirectory();
        var (name, bytes) = PackageFixture.ModuleEntry("modReport.bas", "SomethingElse");
        var zip = PackageFixture.WriteRawZip(temp.PathFor("bad.zip"),
        [
            PackageFixture.ManifestEntry(PackageFixture.SingleModuleManifest(name, bytes)),
            (name, bytes)
        ]);

        var result = new VbaPackageReader().Read(zip, "1.0.0.1", temp.PathFor("extract"));

        Assert.False(result.IsValid);
        Assert.Contains("VB_Name", result.Error);
    }

    [Fact]
    public void Rejects_reserved_input_identifier_in_any_code_file()
    {
        using var temp = new TestDirectory();
        var (name, bytes) = PackageFixture.ModuleEntry("modReport.bas", "modReport", "Private Function Parse(ByVal Input As String) As String\r\nEnd Function\r\n");
        var zip = PackageFixture.WriteRawZip(temp.PathFor("bad.zip"),
        [
            PackageFixture.ManifestEntry(PackageFixture.SingleModuleManifest(name, bytes)),
            (name, bytes)
        ]);

        var result = new VbaPackageReader().Read(zip, "1.0.0.1", temp.PathFor("extract"));

        Assert.False(result.IsValid);
        Assert.Contains("Input", result.Error);
    }

    [Fact]
    public void Rejects_oversized_zip_before_opening_it()
    {
        using var temp = new TestDirectory();
        var zip = temp.PathFor("huge.zip");
        using (var stream = File.Create(zip))
            stream.SetLength(VbaPackageReader.MaxZipBytes + 1);

        var result = new VbaPackageReader().Read(zip, "1.0.0.1", temp.PathFor("extract"));

        Assert.False(result.IsValid);
        Assert.Contains("dung lượng", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Refuses_to_extract_into_non_empty_directory()
    {
        using var temp = new TestDirectory();
        var zip = PackageFixture.BuildPackage(temp);
        var extract = Directory.CreateDirectory(temp.PathFor("extract")).FullName;
        File.WriteAllText(Path.Combine(extract, "keep.txt"), "user file");

        var result = new VbaPackageReader().Read(zip, "1.0.0.1", extract);

        Assert.False(result.IsValid);
        Assert.True(File.Exists(Path.Combine(extract, "keep.txt")));
    }

    [Fact]
    public void Verify_detects_extracted_file_changed_after_reading()
    {
        using var temp = new TestDirectory();
        var zip = PackageFixture.BuildPackage(temp);
        var reader = new VbaPackageReader();
        var package = reader.Read(zip, "1.0.0.1", temp.PathFor("extract")).Package!;
        File.AppendAllText(package.Components[0].FilePath, "' tampered");

        var verification = reader.Verify(package);

        Assert.False(verification.IsValid);
    }
}
