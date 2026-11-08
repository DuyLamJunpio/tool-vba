using System.Net.Http;
using ToolVba.Core;

namespace ToolVba.App.Services;

/// <summary>Composition root: creates the long-lived services once for the whole app.</summary>
internal sealed class AppServices : IDisposable
{
    private static readonly TimeSpan HttpTimeout = TimeSpan.FromSeconds(30);

    private AppServices(HttpClient http, VbaPackageService? packages, string? configurationError)
    {
        Http = http;
        Packages = packages;
        ConfigurationError = configurationError;
    }

    public HttpClient Http { get; }

    /// <summary>Null when the build has no API URL or release key (development builds).</summary>
    public VbaPackageService? Packages { get; }

    public string? ConfigurationError { get; }

    public VbaPackageInstallService Installer { get; } = new(auditLog: new SafeOperationAuditLog());

    public ExcelTrustSettings Trust { get; } = new();

    public AppUpdater Updater { get; } = new(AppConfig.AppUpdateFeedUrl);

    public static AppServices Create()
    {
        var http = new HttpClient { Timeout = HttpTimeout };
        http.DefaultRequestHeaders.UserAgent.ParseAdd($"tool-vba/{AppConfig.AppVersionText}");
        if (!AppConfig.IsConfigured)
            return new(http, null, "Bản này chưa được cấu hình máy chủ cập nhật (release.props).");

        try
        {
            var verifier = new ReleaseSignatureVerifier(AppConfig.ReleasePublicKeyPem!);
            var packages = new VbaPackageService(http, AppConfig.ApiBaseUrl!, verifier, AppConfig.VbaCacheDirectory, AppConfig.AppVersion);
            return new(http, packages, null);
        }
        catch (ArgumentException)
        {
            return new(http, null, "Khóa kiểm tra chữ ký trong bản build không hợp lệ. Hãy liên hệ quản trị viên.");
        }
    }

    public void Dispose() => Http.Dispose();
}
