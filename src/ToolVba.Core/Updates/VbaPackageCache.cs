using System.Text.Json;

namespace ToolVba.Core;

internal sealed record CachedVbaPackage(VbaPackageRelease Release, VbaPackage Package);

/// <summary>
/// Disk cache of verified packages under <c>{root}/{version}/package.zip</c> plus <c>{root}/state.json</c>.
/// Every load re-checks the zip hash and the release signature, so a tampered cache is never used.
/// </summary>
internal sealed class VbaPackageCache(string root, ReleaseSignatureVerifier verifier, VbaPackageReader reader)
{
    private const string StateFileName = "state.json";
    private const string ZipFileName = "package.zip";
    private const int VersionsToKeep = 2;
    private static readonly TimeSpan StaleAfter = TimeSpan.FromHours(1);
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    /// <summary>Why the last <see cref="Load"/> or <see cref="Store"/> could not open a package.</summary>
    public string? LastError { get; private set; }

    public CachedVbaPackage? Load()
    {
        try
        {
            var statePath = Path.Combine(root, StateFileName);
            if (!File.Exists(statePath)) return null;
            var release = JsonSerializer.Deserialize<CacheState>(File.ReadAllText(statePath), JsonOptions)?.Release;
            if (release is null || !IsSafeVersion(release.Version)) return null;
            var zip = Path.Combine(root, release.Version, ZipFileName);
            return File.Exists(zip) ? Open(release, zip) : null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            return null;
        }
    }

    /// <summary>Moves a downloaded, hash-checked zip into the cache and makes it current.</summary>
    public (CachedVbaPackage? Package, string? Error) Store(VbaPackageRelease release, string downloadedZip, string? previousVersion)
    {
        if (!IsSafeVersion(release.Version)) return (null, "Phiên bản gói code VBA không hợp lệ.");
        var directory = Directory.CreateDirectory(Path.Combine(root, release.Version)).FullName;
        var zip = Path.Combine(directory, ZipFileName);
        File.Move(downloadedZip, zip, overwrite: true);
        var opened = Open(release, zip);
        if (opened is null) return (null, LastError ?? "Gói code VBA tải về không hợp lệ.");

        var stateTemp = Path.Combine(root, StateFileName + ".tmp");
        File.WriteAllText(stateTemp, JsonSerializer.Serialize(new CacheState(release with { DownloadUrl = null }), JsonOptions));
        File.Move(stateTemp, Path.Combine(root, StateFileName), overwrite: true);
        PruneVersions(release.Version, previousVersion);
        return (opened, null);
    }

    public string NewDownloadPath()
    {
        Directory.CreateDirectory(root);
        return Path.Combine(root, $"download-{Guid.NewGuid():N}.tmp");
    }

    internal static bool IsSafeVersion(string? version) =>
        System.Version.TryParse(version, out _) && version!.All(c => char.IsAsciiDigit(c) || c == '.');

    private CachedVbaPackage? Open(VbaPackageRelease release, string zip)
    {
        if (!string.Equals(HashFile(zip), release.Sha256, StringComparison.OrdinalIgnoreCase) ||
            !verifier.Verify(ReleaseSignatureVerifier.VbaPackagePayload(release.Version, release.Sha256, release.MinAppVersion), release.Signature))
        {
            LastError = "Gói code VBA đã lưu trên máy không còn khớp chữ ký.";
            return null;
        }

        var versionDirectory = Path.GetDirectoryName(zip)!;
        var files = Path.Combine(versionDirectory, $"files-{Guid.NewGuid():N}");
        var read = reader.Read(zip, release.Version, files);
        if (!read.IsValid)
        {
            LastError = read.Error;
            return null;
        }
        foreach (var old in Directory.EnumerateDirectories(versionDirectory, "files-*").Where(d => !PathEquals(d, files) && IsStale(d)))
            TryDeleteDirectory(old);
        return new(release, read.Package!);
    }

    /// <summary>
    /// Another running copy of the app may still be using a folder it extracted, so only folders
    /// older than <see cref="StaleAfter"/> are removed.
    /// </summary>
    private static bool IsStale(string directory) =>
        DateTime.UtcNow - Directory.GetCreationTimeUtc(directory) > StaleAfter;

    private void PruneVersions(string current, string? previous)
    {
        var keep = new[] { current, previous }.OfType<string>().Take(VersionsToKeep).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var directory in Directory.EnumerateDirectories(root))
        {
            if (!keep.Contains(Path.GetFileName(directory)) && IsSafeVersion(Path.GetFileName(directory)) &&
                DateTime.UtcNow - Directory.GetLastWriteTimeUtc(directory) > StaleAfter)
                TryDeleteDirectory(directory);
        }
        foreach (var stale in Directory.EnumerateFiles(root, "download-*.tmp"))
        {
            try { File.Delete(stale); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static string HashFile(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(stream));
    }

    private static bool PathEquals(string a, string b) =>
        string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);

    private static void TryDeleteDirectory(string path)
    {
        // Best effort: a folder still in use by Excel is removed on a later start.
        try { Directory.Delete(path, recursive: true); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private sealed record CacheState(VbaPackageRelease Release);
}
