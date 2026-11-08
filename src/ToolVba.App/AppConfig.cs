using System.IO;
using System.Reflection;

namespace ToolVba.App;

/// <summary>Build-time release settings (see release.props). None of these values are secrets.</summary>
internal static class AppConfig
{
    private static readonly Assembly Entry = typeof(AppConfig).Assembly;

    public static Version AppVersion { get; } = Entry.GetName().Version ?? new Version(0, 0, 0);

    public static string AppVersionText => $"{AppVersion.Major}.{AppVersion.Minor}.{Math.Max(0, AppVersion.Build)}";

    /// <summary>
    /// https only: the admin password, the access token and the app update feed travel over this URL.
    /// Plain http is accepted for a localhost API during development.
    /// </summary>
    public static Uri? ApiBaseUrl { get; } =
        Uri.TryCreate(Metadata("ToolVbaApiBaseUrl"), UriKind.Absolute, out var url) &&
        (url.Scheme == Uri.UriSchemeHttps || url.Scheme == Uri.UriSchemeHttp && url.IsLoopback)
            ? new Uri(url.AbsoluteUri.EndsWith('/') ? url.AbsoluteUri : url.AbsoluteUri + "/")
            : null;

    public static string? ReleasePublicKeyPem { get; } = ToPem(Metadata("ToolVbaReleasePublicKey"));

    public static bool IsConfigured => ApiBaseUrl is not null && ReleasePublicKeyPem is not null;

    public static string? AppUpdateFeedUrl => ApiBaseUrl is null ? null : new Uri(ApiBaseUrl, "api/v1/tool/app/").AbsoluteUri;

    public static string VbaCacheDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ToolVba", "vba");

    private static string? Metadata(string key) =>
        Entry.GetCustomAttributes<AssemblyMetadataAttribute>().FirstOrDefault(a => a.Key == key)?.Value is { Length: > 0 } value
            ? value.Trim()
            : null;

    private static string? ToPem(string? base64Spki)
    {
        if (base64Spki is null) return null;
        if (base64Spki.Contains("BEGIN PUBLIC KEY", StringComparison.Ordinal)) return base64Spki;
        return $"-----BEGIN PUBLIC KEY-----\n{base64Spki}\n-----END PUBLIC KEY-----";
    }
}
