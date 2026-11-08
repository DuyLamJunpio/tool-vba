using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;

namespace ToolVba.Core;

/// <summary>
/// Keeps the company VBA package current: asks the API for the latest signed release, downloads and
/// verifies it, and falls back to the last verified package when the server cannot be reached.
/// </summary>
public sealed class VbaPackageService
{
    private const string LatestPath = "api/v1/tool/vba-packages/latest";
    private const int MaxReleaseJsonBytes = 64 * 1024;
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    private readonly HttpClient _http;
    private readonly Uri _latestUri;
    private readonly bool _allowHttp;
    private readonly ReleaseSignatureVerifier _verifier;
    private readonly Version _appVersion;
    private readonly VbaPackageCache _cache;

    public VbaPackageService(HttpClient http, Uri apiBaseUrl, ReleaseSignatureVerifier verifier,
        string cacheRoot, Version appVersion, VbaPackageReader? reader = null)
    {
        _http = http;
        _latestUri = new Uri(EnsureTrailingSlash(apiBaseUrl), LatestPath);
        // Plain http is only acceptable against a local development API.
        _allowHttp = apiBaseUrl.Scheme == Uri.UriSchemeHttp && apiBaseUrl.IsLoopback;
        _verifier = verifier;
        _appVersion = Normalize(appVersion);
        _cache = new VbaPackageCache(cacheRoot, verifier, reader ?? new VbaPackageReader());
    }

    public async Task<VbaPackageStatus> RefreshAsync(CancellationToken cancellationToken = default)
    {
        var cached = _cache.Load();
        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, _latestUri);
            if (cached is not null)
                request.Headers.IfNoneMatch.Add(new EntityTagHeaderValue($"\"{cached.Release.Sha256}\""));
            using var response = await _http.SendAsync(request, cancellationToken);

            if (response.StatusCode == HttpStatusCode.NotModified && cached is not null)
                return Ready(VbaPackageAvailability.Latest, cached);
            if (response.StatusCode == HttpStatusCode.NotFound)
                return Fallback(cached, "Chưa có gói code VBA nào được xuất bản. Hãy liên hệ quản trị viên.");
            if (!response.IsSuccessStatusCode)
                return Fallback(cached, $"Máy chủ cập nhật đang lỗi ({(int)response.StatusCode}).");

            var release = await ReadReleaseAsync(response, cancellationToken);
            if (release is null)
                return Fallback(cached, "Máy chủ trả thông tin gói code VBA không hợp lệ.");
            if (!_verifier.Verify(ReleaseSignatureVerifier.VbaPackagePayload(release.Version, release.Sha256, release.MinAppVersion), release.Signature))
                return Fallback(cached, "Gói code VBA mới không có chữ ký hợp lệ nên đã bị bỏ qua.");
            if (cached is not null && cached.Release.Version == release.Version &&
                string.Equals(cached.Release.Sha256, release.Sha256, StringComparison.OrdinalIgnoreCase))
                return Ready(VbaPackageAvailability.Latest, cached with { Release = release });

            var (downloaded, error) = await DownloadAsync(release, cached?.Release.Version, cancellationToken);
            return downloaded is null ? Fallback(cached, error!) : Ready(VbaPackageAvailability.Latest, downloaded);
        }
        catch (Exception ex) when (IsConnectivityFailure(ex, cancellationToken))
        {
            return Fallback(cached, null);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Fallback(cached, "Không ghi được gói code VBA vào máy (%LocalAppData%\\ToolVba).");
        }
    }

    private async Task<(CachedVbaPackage? Package, string? Error)> DownloadAsync(
        VbaPackageRelease release, string? previousVersion, CancellationToken cancellationToken)
    {
        var temp = _cache.NewDownloadPath();
        try
        {
            using (var response = await _http.GetAsync(release.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellationToken))
            {
                if (!response.IsSuccessStatusCode)
                    return (null, "Không tải được gói code VBA mới từ máy chủ.");
                if (response.Content.Headers.ContentLength > VbaPackageReader.MaxZipBytes)
                    return (null, "Gói code VBA mới vượt quá dung lượng cho phép.");
                await using var source = await response.Content.ReadAsStreamAsync(cancellationToken);
                await using var target = File.Create(temp);
                if (!await CopyBoundedAsync(source, target, VbaPackageReader.MaxZipBytes, cancellationToken))
                    return (null, "Gói code VBA mới vượt quá dung lượng cho phép.");
            }

            await using (var stream = File.OpenRead(temp))
            {
                var actual = Convert.ToHexStringLower(await System.Security.Cryptography.SHA256.HashDataAsync(stream, cancellationToken));
                if (!string.Equals(actual, release.Sha256, StringComparison.OrdinalIgnoreCase))
                    return (null, "File tải về không khớp SHA-256 đã ký. Gói code VBA mới đã bị bỏ qua.");
            }
            return _cache.Store(release, temp, previousVersion);
        }
        finally
        {
            try { if (File.Exists(temp)) File.Delete(temp); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private async Task<VbaPackageRelease?> ReadReleaseAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        try
        {
            await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
            using var buffer = new MemoryStream();
            if (!await CopyBoundedAsync(stream, buffer, MaxReleaseJsonBytes, cancellationToken)) return null;
            using var document = JsonDocument.Parse(buffer.ToArray());
            if (document.RootElement.ValueKind != JsonValueKind.Object ||
                !document.RootElement.TryGetProperty("data", out var data) || data.ValueKind != JsonValueKind.Object)
                return null;
            var release = data.Deserialize<VbaPackageRelease>(JsonOptions);
            return IsWellFormed(release) ? release : null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private bool IsWellFormed(VbaPackageRelease? release) =>
        release is not null &&
        VbaPackageCache.IsSafeVersion(release.Version) &&
        release.Sha256 is { Length: 64 } && release.Sha256.All(Uri.IsHexDigit) &&
        release.SizeBytes is > 0 and <= VbaPackageReader.MaxZipBytes &&
        !string.IsNullOrWhiteSpace(release.Signature) &&
        (release.MinAppVersion is null || Version.TryParse(release.MinAppVersion, out _)) &&
        Uri.TryCreate(release.DownloadUrl, UriKind.Absolute, out var url) &&
        (url.Scheme == Uri.UriSchemeHttps || _allowHttp && url.Scheme == Uri.UriSchemeHttp);

    private VbaPackageStatus Ready(VbaPackageAvailability availability, CachedVbaPackage cached, string? prefix = null)
    {
        var version = cached.Release.Version;
        if (Version.TryParse(cached.Release.MinAppVersion, out var required) && Normalize(required) > _appVersion)
        {
            return new(availability, cached.Package, cached.Release,
                $"Code VBA v{version} cần ứng dụng bản {cached.Release.MinAppVersion} trở lên. Hãy bấm Cập nhật ứng dụng.",
                cached.Release.MinAppVersion);
        }
        var message = availability == VbaPackageAvailability.Latest
            ? $"Code VBA v{version} — mới nhất"
            : $"{prefix ?? "Không kết nối được máy chủ."} Đang dùng code VBA đã lưu v{version}.";
        return new(availability, cached.Package, cached.Release, message);
    }

    private VbaPackageStatus Fallback(CachedVbaPackage? cached, string? reason)
    {
        if (cached is not null) return Ready(VbaPackageAvailability.CachedOffline, cached, reason);
        return VbaPackageStatus.Unavailable(reason ??
            "Cần kết nối mạng ở lần dùng đầu tiên để tải code VBA. Hãy kiểm tra mạng rồi thử lại.");
    }

    private static async Task<bool> CopyBoundedAsync(Stream source, Stream target, long maxBytes, CancellationToken cancellationToken)
    {
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = await source.ReadAsync(buffer, cancellationToken)) > 0)
        {
            total += read;
            if (total > maxBytes) return false;
            await target.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
        }
        return true;
    }

    /// <summary>Network failures only; disk errors are reported separately so the message stays accurate.</summary>
    private static bool IsConnectivityFailure(Exception ex, CancellationToken cancellationToken) =>
        ex is HttpRequestException or HttpIOException ||
        ex is TaskCanceledException && !cancellationToken.IsCancellationRequested;

    private static Uri EnsureTrailingSlash(Uri uri) =>
        uri.AbsoluteUri.EndsWith('/') ? uri : new Uri(uri.AbsoluteUri + "/");

    internal static Version Normalize(Version version) =>
        new(version.Major, Math.Max(0, version.Minor), Math.Max(0, version.Build));
}
