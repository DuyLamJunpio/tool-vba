using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace ToolVba.Core;

/// <summary>
/// Reads a company VBA package zip (docs/tool-release-api-contract.md §1.1), validates every entry
/// before anything is written, and extracts only the referenced files into an empty directory.
/// </summary>
public sealed class VbaPackageReader
{
    public const long MaxZipBytes = 5L * 1024 * 1024;
    public const long MaxExtractedBytes = 20L * 1024 * 1024;
    public const int MaxEntries = 200;
    private const long MaxManifestBytes = 256 * 1024;

    private static readonly Regex SafeEntryName = new(@"^[A-Za-z0-9_.-]{1,64}$", RegexOptions.CultureInvariant);

    // Windows device names cannot be created as files ("NUL.bas" would fail on every machine).
    private static readonly Regex ReservedDeviceName = new(@"^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    public VbaPackageReadResult Read(string zipPath, string version, string extractDirectory)
    {
        try
        {
            var info = new FileInfo(zipPath);
            if (!info.Exists) return VbaPackageReadResult.Invalid("Không tìm thấy file gói code VBA.");
            if (info.Length > MaxZipBytes) return VbaPackageReadResult.Invalid("Gói code VBA vượt quá dung lượng cho phép (5 MB).");
            if (Directory.Exists(extractDirectory) && Directory.EnumerateFileSystemEntries(extractDirectory).Any())
                return VbaPackageReadResult.Invalid("Thư mục giải nén gói code VBA phải trống.");

            using var archive = ZipFile.OpenRead(zipPath);
            var (entries, entryError) = IndexEntries(archive);
            if (entryError is not null) return VbaPackageReadResult.Invalid(entryError);
            var (manifest, manifestError) = ReadManifest(entries);
            if (manifestError is not null) return VbaPackageReadResult.Invalid(manifestError);
            var validationError = ValidateManifest(manifest!, entries);
            if (validationError is not null) return VbaPackageReadResult.Invalid(validationError);

            var contents = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var component in manifest!.Components)
            {
                var bytes = ReadEntry(entries[component.File]);
                var error = CheckHash(component.File, bytes, component.Sha256) ?? CheckSource(component, bytes);
                if (error is not null) return VbaPackageReadResult.Invalid(error);
                contents[component.File] = bytes;
                if (component.ExtraFile is null) continue;
                var extra = ReadEntry(entries[component.ExtraFile]);
                error = CheckHash(component.ExtraFile, extra, component.ExtraSha256!);
                if (error is not null) return VbaPackageReadResult.Invalid(error);
                contents[component.ExtraFile] = extra;
            }

            return VbaPackageReadResult.Valid(Extract(manifest, contents, version, extractDirectory));
        }
        catch (InvalidDataException)
        {
            return VbaPackageReadResult.Invalid("Gói code VBA không phải file zip hợp lệ.");
        }
        catch (JsonException)
        {
            return VbaPackageReadResult.Invalid("package.json trong gói code VBA không hợp lệ.");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return VbaPackageReadResult.Invalid("Không thể đọc hoặc giải nén gói code VBA. Kiểm tra quyền truy cập thư mục.");
        }
    }

    /// <summary>Re-hashes extracted files so callers can detect changes after <see cref="Read"/>.</summary>
    public VbaPackageReadResult Verify(VbaPackage package)
    {
        try
        {
            foreach (var component in package.Components)
            {
                if (!HashMatches(component.FilePath, component.Sha256) ||
                    component.ExtraFilePath is not null && !HashMatches(component.ExtraFilePath, component.ExtraSha256!))
                {
                    return VbaPackageReadResult.Invalid($"File của component {component.Name} đã thay đổi sau khi kiểm tra. Hãy mở lại ứng dụng để tải lại gói code VBA.");
                }
            }
            return VbaPackageReadResult.Valid(package);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return VbaPackageReadResult.Invalid("Không đọc được file của gói code VBA đã lưu.");
        }
    }

    internal static string Sha256Hex(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

    private static bool HashMatches(string path, string expected)
    {
        using var stream = File.OpenRead(path);
        return string.Equals(Convert.ToHexStringLower(SHA256.HashData(stream)), expected, StringComparison.OrdinalIgnoreCase);
    }

    private static (Dictionary<string, ZipArchiveEntry> Entries, string? Error) IndexEntries(ZipArchive archive)
    {
        var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        if (archive.Entries.Count > MaxEntries)
            return (entries, $"Gói code VBA có quá nhiều file (tối đa {MaxEntries}).");
        long total = 0;
        foreach (var entry in archive.Entries)
        {
            if (!IsSafeEntryName(entry.FullName))
                return (entries, $"Gói code VBA có tên file không hợp lệ: {entry.FullName}");
            if (!entries.TryAdd(entry.FullName, entry))
                return (entries, $"Gói code VBA có tên file trùng: {entry.FullName}");
            total += entry.Length;
            if (total > MaxExtractedBytes)
                return (entries, "Gói code VBA vượt quá dung lượng giải nén cho phép (20 MB).");
        }
        return (entries, null);
    }

    private static bool IsSafeEntryName(string name) =>
        SafeEntryName.IsMatch(name) && !name.Contains("..", StringComparison.Ordinal) && name.Trim('.').Length > 0 &&
        !name.EndsWith('.') && !ReservedDeviceName.IsMatch(name);

    private static (VbaPackageManifest? Manifest, string? Error) ReadManifest(Dictionary<string, ZipArchiveEntry> entries)
    {
        if (!entries.TryGetValue(VbaPackageManifest.FileName, out var entry))
            return (null, "Gói code VBA thiếu package.json.");
        if (entry.Length > MaxManifestBytes)
            return (null, "package.json trong gói code VBA quá lớn.");
        var manifest = JsonSerializer.Deserialize<VbaPackageManifest>(ReadEntry(entry), VbaPackageManifest.JsonOptions);
        return manifest is null ? (null, "package.json trong gói code VBA không hợp lệ.") : (manifest, null);
    }

    private static string? ValidateManifest(VbaPackageManifest manifest, Dictionary<string, ZipArchiveEntry> entries)
    {
        if (manifest.FormatVersion != VbaPackageManifest.CurrentFormatVersion)
            return $"Gói code VBA dùng định dạng {manifest.FormatVersion}, ứng dụng này chưa hỗ trợ. Hãy cập nhật ứng dụng.";
        if (manifest.Components is null || manifest.Components.Count == 0)
            return "Gói code VBA không có component nào.";
        if (manifest.Components.Count(c => c?.Kind == VbaComponentKind.Workbook) > 1)
            return "Gói code VBA chỉ được có một component ThisWorkbook.";

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var component in manifest.Components)
        {
            var error = ValidateComponent(component, entries);
            if (error is not null) return error;
            if (!names.Add(component.Name))
                return $"Gói code VBA có tên component trùng: {component.Name}.";
            if (!files.Add(component.File) || component.ExtraFile is not null && !files.Add(component.ExtraFile))
                return $"Gói code VBA dùng một file cho nhiều component: {component.Name}.";
        }
        return null;
    }

    private static string? ValidateComponent(VbaPackageManifestComponent? component, Dictionary<string, ZipArchiveEntry> entries)
    {
        if (component is null || !VbaSourceRules.IsValidComponentName(component.Name))
            return "Gói code VBA có component thiếu tên hoặc tên không hợp lệ.";
        var extension = VbaSourceRules.ExtensionFor(component.Kind);
        if (extension is null)
            return $"Component {component.Name} có loại chưa được hỗ trợ.";
        if (string.IsNullOrWhiteSpace(component.File) || !Path.GetExtension(component.File).Equals(extension, StringComparison.OrdinalIgnoreCase))
            return $"Component {component.Name} có phần mở rộng file không đúng với loại {component.Kind}.";
        if (!entries.ContainsKey(component.File))
            return $"Gói code VBA thiếu file {component.File}.";
        if (!IsSha256(component.Sha256))
            return $"SHA-256 của {component.File} trong package.json không hợp lệ.";
        if (component.SheetName is not null &&
            (component.Kind != VbaComponentKind.Sheet || component.SheetName.Length is 0 or > VbaSourceRules.MaxSheetNameLength))
            return $"Component {component.Name} có sheetName không hợp lệ.";

        if (component.Kind != VbaComponentKind.Form)
            return component.ExtraFile is null ? null : $"Component {component.Name} không được có file đi kèm.";
        if (component.ExtraFile is null || !IsSha256(component.ExtraSha256) ||
            !component.ExtraFile.Equals(Path.ChangeExtension(component.File, ".frx"), StringComparison.OrdinalIgnoreCase))
            return $"UserForm {component.Name} thiếu file .frx đi kèm cùng tên.";
        return entries.ContainsKey(component.ExtraFile) ? null : $"Gói code VBA thiếu file {component.ExtraFile}.";
    }

    private static bool IsSha256(string? value) => value is { Length: 64 } && value.All(Uri.IsHexDigit);

    private static string? CheckHash(string file, byte[] bytes, string expected) =>
        string.Equals(Sha256Hex(bytes), expected, StringComparison.OrdinalIgnoreCase)
            ? null
            : $"SHA-256 của {file} không khớp package.json. Gói code VBA có thể đã bị sửa.";

    private static string? CheckSource(VbaPackageManifestComponent component, byte[] bytes)
    {
        // Exported .bas/.cls/.frm files use the ANSI code page; document code is stored as UTF-8.
        var encoding = component.Kind is VbaComponentKind.Workbook or VbaComponentKind.Sheet ? Encoding.UTF8 : Encoding.Latin1;
        var lines = VbaSourceRules.SplitLines(encoding.GetString(bytes));
        var reserved = VbaSourceRules.FindReservedInputDeclaration(lines);
        if (reserved is not null) return $"{component.Name}: {reserved}";
        if (component.Kind is VbaComponentKind.Workbook or VbaComponentKind.Sheet) return null;
        // A non-ASCII VB_Name depends on the ANSI code page of the exporting machine, so it cannot be
        // compared reliably here; the installer still checks the name Excel assigns on import.
        if (!component.Name.All(char.IsAscii)) return null;
        var vbName = VbaSourceRules.ReadVbName(lines);
        return string.Equals(vbName, component.Name, StringComparison.OrdinalIgnoreCase)
            ? null
            : $"File {component.File} có Attribute VB_Name khác tên component {component.Name}.";
    }

    private static byte[] ReadEntry(ZipArchiveEntry entry)
    {
        if (entry.Length > MaxExtractedBytes)
            throw new InvalidDataException("Entry is larger than allowed.");
        using var source = entry.Open();
        using var buffer = new MemoryStream((int)entry.Length);
        var chunk = new byte[81920];
        int read;
        while ((read = source.Read(chunk, 0, chunk.Length)) > 0)
        {
            if (buffer.Length + read > entry.Length)
                throw new InvalidDataException("Entry is larger than its declared size.");
            buffer.Write(chunk, 0, read);
        }
        return buffer.ToArray();
    }

    private static VbaPackage Extract(VbaPackageManifest manifest, Dictionary<string, byte[]> contents, string version, string extractDirectory)
    {
        var root = Directory.CreateDirectory(extractDirectory).FullName;
        string Write(string name)
        {
            var target = Path.GetFullPath(Path.Combine(root, name));
            if (!target.StartsWith(root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Entry escapes the extraction directory.");
            File.WriteAllBytes(target, contents[name]);
            return target;
        }

        var files = manifest.Components.Select(c => new VbaPackageComponentFile(
            c.Name, c.Kind, Write(c.File), c.Sha256.ToLowerInvariant(),
            c.ExtraFile is null ? null : Write(c.ExtraFile), c.ExtraSha256?.ToLowerInvariant(), c.SheetName)).ToList();
        return new VbaPackage(version, root, files);
    }
}
