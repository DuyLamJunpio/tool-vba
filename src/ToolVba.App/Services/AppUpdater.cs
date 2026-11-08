using System.Net.Http;
using Velopack;

namespace ToolVba.App.Services;

internal sealed record AppUpdateCheck(UpdateInfo? Update, string? Version, string? Notes, string? Error);

/// <summary>Thin wrapper over Velopack so view models do not depend on its types directly.</summary>
internal sealed class AppUpdater
{
    private readonly UpdateManager? _manager;

    public AppUpdater(string? feedUrl)
    {
        if (feedUrl is not null) _manager = new UpdateManager(feedUrl);
    }

    /// <summary>False for development runs and for copies not installed with Setup.exe.</summary>
    public bool IsSupported => _manager?.IsInstalled == true;

    public async Task<AppUpdateCheck> CheckAsync()
    {
        if (!IsSupported) return new(null, null, null, null);
        try
        {
            var update = await _manager!.CheckForUpdatesAsync();
            return update is null
                ? new(null, null, null, null)
                : new(update, update.TargetFullRelease.Version.ToString(), update.TargetFullRelease.NotesMarkdown, null);
        }
        catch (Exception ex)
        {
            // Boundary to Velopack: a bad feed or network error must never stop the app from starting.
            AppErrorLog.Write("update-check", ex);
            return new(null, null, null, "Không kiểm tra được bản cập nhật ứng dụng.");
        }
    }

    public Task DownloadAsync(UpdateInfo update, Action<int> progress, CancellationToken cancellationToken) =>
        _manager!.DownloadUpdatesAsync(update, progress, cancellationToken);

    public void ApplyAndRestart(UpdateInfo update) => _manager!.ApplyUpdatesAndRestart(update);
}
