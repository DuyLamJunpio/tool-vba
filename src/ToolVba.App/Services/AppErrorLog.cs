using System.IO;
using System.Text;
using System.Text.Json;

namespace ToolVba.App.Services;

/// <summary>
/// Local record of unexpected errors for support. Only the context, exception type and HRESULT are written:
/// messages and stack traces can contain file paths, workbook data or VBA code.
/// </summary>
internal static class AppErrorLog
{
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ToolVba", "logs", "app-errors.jsonl");
    private static readonly object Gate = new();

    public static void Write(string context, Exception exception)
    {
        var line = JsonSerializer.Serialize(new
        {
            timestampUtc = DateTimeOffset.UtcNow,
            context,
            type = exception.GetType().FullName,
            hresult = $"0x{unchecked((uint)exception.HResult):X8}",
            inner = exception.InnerException?.GetType().FullName,
            appVersion = AppConfig.AppVersionText
        });
        try
        {
            lock (Gate)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
                File.AppendAllText(LogPath, line + Environment.NewLine, new UTF8Encoding(false));
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // Logging must never take the app down; the user already sees a message for the failure itself.
        }
    }
}
