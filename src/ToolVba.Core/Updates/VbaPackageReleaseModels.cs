namespace ToolVba.Core;

/// <summary>Release metadata returned by <c>GET /api/v1/tool/vba-packages/latest</c> (contract §1.3).</summary>
public sealed record VbaPackageRelease(
    string Version,
    string Sha256,
    long SizeBytes,
    string? MinAppVersion,
    string? ReleaseNotes,
    DateTimeOffset? PublishedAt,
    string Signature,
    string? DownloadUrl);

public enum VbaPackageAvailability
{
    /// <summary>The package matches what the server currently publishes.</summary>
    Latest,

    /// <summary>The server could not be used; a previously verified package is used instead.</summary>
    CachedOffline,

    /// <summary>No verified package is available on this machine.</summary>
    Unavailable
}

public sealed record VbaPackageStatus(
    VbaPackageAvailability Availability,
    VbaPackage? Package,
    VbaPackageRelease? Release,
    string Message,
    string? RequiredAppVersion = null)
{
    public bool CanInstall => Package is not null && RequiredAppVersion is null;

    public static VbaPackageStatus Unavailable(string message) => new(VbaPackageAvailability.Unavailable, null, null, message);
}
