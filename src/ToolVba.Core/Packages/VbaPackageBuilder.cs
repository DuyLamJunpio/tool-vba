using System.IO.Compression;
using System.Text.Json;

namespace ToolVba.Core;

/// <summary>Builds a package zip from staged component files and checks it with <see cref="VbaPackageReader"/>.</summary>
public static class VbaPackageBuilder
{
    public static VbaPackageBuildResult Build(IReadOnlyList<VbaPackageSource> sources, string zipPath)
    {
        if (sources.Count == 0)
            return VbaPackageBuildResult.Invalid("Không có component nào để đóng gói.");

        var components = new List<VbaPackageManifestComponent>();
        foreach (var source in sources)
        {
            var (component, error) = Describe(source);
            if (error is not null) return VbaPackageBuildResult.Invalid(error);
            components.Add(component!);
        }

        var tempZip = zipPath + ".tmp";
        var checkDirectory = Path.Combine(Path.GetTempPath(), $"tool-vba-check-{Guid.NewGuid():N}");
        try
        {
            WriteZip(tempZip, sources, components);
            var check = new VbaPackageReader().Read(tempZip, "0.0.0.0", checkDirectory);
            if (!check.IsValid) return VbaPackageBuildResult.Invalid(check.Error!);
            File.Move(tempZip, zipPath, overwrite: true);
            return new(true, null, components);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return VbaPackageBuildResult.Invalid("Không thể ghi file gói code VBA. Kiểm tra quyền truy cập thư mục.");
        }
        finally
        {
            TryDelete(tempZip);
            try { if (Directory.Exists(checkDirectory)) Directory.Delete(checkDirectory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static (VbaPackageManifestComponent? Component, string? Error) Describe(VbaPackageSource source)
    {
        if (!VbaSourceRules.IsValidComponentName(source.Name))
            return (null, $"Tên component không hợp lệ: {source.Name}.");
        if (!File.Exists(source.FilePath))
            return (null, $"Không tìm thấy file của component {source.Name}.");
        var file = Path.GetFileName(source.FilePath);
        string? extra = null;
        string? extraHash = null;
        if (source.Kind == VbaComponentKind.Form)
        {
            if (source.ExtraFilePath is null || !File.Exists(source.ExtraFilePath))
                return (null, $"UserForm {source.Name} thiếu file .frx đi kèm.");
            extra = Path.GetFileName(source.ExtraFilePath);
            extraHash = VbaPackageReader.Sha256Hex(File.ReadAllBytes(source.ExtraFilePath));
        }
        var sheetName = source.Kind == VbaComponentKind.Sheet ? source.SheetName : null;
        return (new(source.Name, source.Kind, file, VbaPackageReader.Sha256Hex(File.ReadAllBytes(source.FilePath)), extra, extraHash, sheetName), null);
    }

    private static void WriteZip(string path, IReadOnlyList<VbaPackageSource> sources, IReadOnlyList<VbaPackageManifestComponent> components)
    {
        using var stream = File.Create(path);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);
        var manifest = new VbaPackageManifest(VbaPackageManifest.CurrentFormatVersion, components);
        var manifestEntry = archive.CreateEntry(VbaPackageManifest.FileName, CompressionLevel.Optimal);
        using (var manifestStream = manifestEntry.Open())
            JsonSerializer.Serialize(manifestStream, manifest, VbaPackageManifest.JsonOptions);

        foreach (var source in sources)
        {
            archive.CreateEntryFromFile(source.FilePath, Path.GetFileName(source.FilePath), CompressionLevel.Optimal);
            if (source.ExtraFilePath is not null)
                archive.CreateEntryFromFile(source.ExtraFilePath, Path.GetFileName(source.ExtraFilePath), CompressionLevel.Optimal);
        }
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
