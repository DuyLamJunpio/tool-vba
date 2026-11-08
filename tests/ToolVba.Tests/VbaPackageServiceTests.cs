using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class VbaPackageServiceTests : IDisposable
{
    private static readonly Uri ApiBase = new("https://api.test/");
    private const string LatestUrl = "https://api.test/api/v1/tool/vba-packages/latest";
    private const string DownloadUrl = "https://storage.test/vba/package.zip";

    private readonly TestDirectory _temp = new();
    private readonly ECDsa _key = ECDsa.Create(ECCurve.NamedCurves.nistP256);
    private readonly FakeHandler _handler = new();

    public void Dispose()
    {
        _key.Dispose();
        _temp.Dispose();
    }

    [Fact]
    public async Task Downloads_verifies_and_caches_latest_package()
    {
        ServeRelease("2026.11.7.1");

        var status = await CreateService().RefreshAsync();

        Assert.Equal(VbaPackageAvailability.Latest, status.Availability);
        Assert.True(status.CanInstall);
        Assert.Equal("2026.11.7.1", status.Package!.Version);
        Assert.Equal(4, status.Package.Components.Count);
        Assert.Contains("2026.11.7.1", status.Message);
    }

    [Fact]
    public async Task Sends_etag_and_uses_cache_on_not_modified_without_downloading_again()
    {
        var release = ServeRelease("2026.11.7.1");
        await CreateService().RefreshAsync();
        _handler.Requests.Clear();
        _handler.Routes[LatestUrl] = request =>
        {
            Assert.Equal($"\"{release.Sha256}\"", request.Headers.IfNoneMatch.Single().Tag);
            return new HttpResponseMessage(HttpStatusCode.NotModified);
        };

        var status = await CreateService().RefreshAsync();

        Assert.Equal(VbaPackageAvailability.Latest, status.Availability);
        Assert.Equal("2026.11.7.1", status.Package!.Version);
        Assert.DoesNotContain(_handler.Requests, uri => uri == DownloadUrl);
    }

    [Fact]
    public async Task Uses_cached_package_when_server_is_unreachable()
    {
        ServeRelease("2026.11.7.1");
        await CreateService().RefreshAsync();
        _handler.ThrowNetworkError = true;

        var status = await CreateService().RefreshAsync();

        Assert.Equal(VbaPackageAvailability.CachedOffline, status.Availability);
        Assert.True(status.CanInstall);
        Assert.Contains("2026.11.7.1", status.Message);
    }

    [Fact]
    public async Task Reports_unavailable_when_offline_on_first_use()
    {
        _handler.ThrowNetworkError = true;

        var status = await CreateService().RefreshAsync();

        Assert.Equal(VbaPackageAvailability.Unavailable, status.Availability);
        Assert.False(status.CanInstall);
        Assert.Contains("kết nối", status.Message);
    }

    [Fact]
    public async Task Rejects_release_with_invalid_signature_and_does_not_download_it()
    {
        ServeRelease("2026.11.7.1", tamperSignature: true);

        var status = await CreateService().RefreshAsync();

        Assert.Equal(VbaPackageAvailability.Unavailable, status.Availability);
        Assert.Contains("chữ ký", status.Message);
        Assert.DoesNotContain(_handler.Requests, uri => uri == DownloadUrl);
    }

    [Fact]
    public async Task Rejects_downloaded_zip_whose_hash_differs_from_signed_release()
    {
        ServeRelease("2026.11.7.1");
        _handler.Routes[DownloadUrl] = _ => Bytes(Encoding.ASCII.GetBytes("not the signed zip"));

        var status = await CreateService().RefreshAsync();

        Assert.False(status.CanInstall);
        Assert.Contains("SHA-256", status.Message);
    }

    [Fact]
    public async Task Blocks_install_when_package_requires_newer_app()
    {
        ServeRelease("2026.11.7.1", minAppVersion: "2.0.0");

        var status = await CreateService(new Version(1, 4, 0)).RefreshAsync();

        Assert.Equal("2.0.0", status.RequiredAppVersion);
        Assert.False(status.CanInstall);
    }

    [Theory]
    [InlineData("null")]
    [InlineData("[]")]
    [InlineData("\"text\"")]
    [InlineData("{\"data\":[1]}")]
    public async Task Treats_unexpected_json_shapes_as_invalid_response_without_throwing(string body)
    {
        _handler.Routes[LatestUrl] = _ => new HttpResponseMessage(HttpStatusCode.OK)
        {
            Content = new StringContent(body, Encoding.UTF8, "application/json")
        };

        var status = await CreateService().RefreshAsync();

        Assert.Equal(VbaPackageAvailability.Unavailable, status.Availability);
        Assert.Contains("không hợp lệ", status.Message);
    }

    [Fact]
    public async Task Reports_when_nothing_has_been_published()
    {
        _handler.Routes[LatestUrl] = _ => new HttpResponseMessage(HttpStatusCode.NotFound);

        var status = await CreateService().RefreshAsync();

        Assert.Equal(VbaPackageAvailability.Unavailable, status.Availability);
        Assert.Contains("Chưa có", status.Message);
    }

    [Fact]
    public async Task Rejects_tampered_cache_when_offline()
    {
        ServeRelease("2026.11.7.1");
        await CreateService().RefreshAsync();
        foreach (var zip in Directory.EnumerateFiles(CacheRoot, "package.zip", SearchOption.AllDirectories))
            File.AppendAllText(zip, "x");
        _handler.ThrowNetworkError = true;

        var status = await CreateService().RefreshAsync();

        Assert.False(status.CanInstall);
    }

    [Fact]
    public async Task Follows_server_rollback_to_an_older_version()
    {
        ServeRelease("2026.11.7.2");
        await CreateService().RefreshAsync();
        ServeRelease("2026.11.7.1", variant: "rollback");

        var status = await CreateService().RefreshAsync();

        Assert.Equal(VbaPackageAvailability.Latest, status.Availability);
        Assert.Equal("2026.11.7.1", status.Package!.Version);
    }

    private string CacheRoot => _temp.PathFor("cache");

    private VbaPackageService CreateService(Version? appVersion = null) =>
        new(new HttpClient(_handler), ApiBase, new ReleaseSignatureVerifier(_key.ExportSubjectPublicKeyInfoPem()),
            CacheRoot, appVersion ?? new Version(1, 0, 0));

    private VbaPackageRelease ServeRelease(string version, string? minAppVersion = null, bool tamperSignature = false, string variant = "")
    {
        var zip = PackageFixture.BuildPackage(_temp, $"pkg-{version}-{variant}-{Guid.NewGuid():N}.zip");
        var bytes = File.ReadAllBytes(zip);
        var sha = PackageFixture.Sha256(bytes);
        var signature = ReleaseSignatureVerifierTests.Sign(_key,
            ReleaseSignatureVerifier.VbaPackagePayload(tamperSignature ? version + "0" : version, sha, minAppVersion));
        var release = new VbaPackageRelease(version, sha, bytes.Length, minAppVersion, "notes",
            DateTimeOffset.Parse("2026-11-07T03:00:00Z"), signature, DownloadUrl);
        var envelope = JsonSerializer.Serialize(new { success = true, statusCode = 200, data = release },
            new JsonSerializerOptions(JsonSerializerDefaults.Web));
        _handler.Routes[LatestUrl] = _ =>
        {
            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(envelope, Encoding.UTF8, "application/json")
            };
            response.Headers.ETag = new($"\"{sha}\"");
            return response;
        };
        _handler.Routes[DownloadUrl] = _ => Bytes(bytes);
        return release;
    }

    private static HttpResponseMessage Bytes(byte[] bytes) =>
        new(HttpStatusCode.OK) { Content = new ByteArrayContent(bytes) };

    private sealed class FakeHandler : HttpMessageHandler
    {
        public Dictionary<string, Func<HttpRequestMessage, HttpResponseMessage>> Routes { get; } = [];
        public List<string> Requests { get; } = [];
        public bool ThrowNetworkError { get; set; }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var uri = request.RequestUri!.ToString();
            Requests.Add(uri);
            if (ThrowNetworkError) throw new HttpRequestException("offline");
            return Task.FromResult(Routes.TryGetValue(uri, out var route)
                ? route(request)
                : new HttpResponseMessage(HttpStatusCode.NotFound));
        }
    }
}
