namespace ToolVba.Core;

internal sealed record InstallPaths(string Input, string Output, bool IsXlsx)
{
    public static (InstallPaths? Paths, string? Error) Resolve(VbaPackageInstallRequest request)
    {
        string input;
        string output;
        try
        {
            input = Path.GetFullPath(request.InputWorkbookPath);
            output = Path.GetFullPath(request.OutputPath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return (null, "Có đường dẫn không hợp lệ hoặc quá dài.");
        }

        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase))
            return (null, "File kết quả không được trùng đường dẫn file nguồn.");
        var extension = Path.GetExtension(input);
        var isXlsx = extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase);
        if (!isXlsx && !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            return (null, "Chỉ hỗ trợ file Excel .xlsx hoặc .xlsm.");
        if (!Path.GetExtension(output).Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
            return (null, "File kết quả phải có phần mở rộng .xlsm.");
        if (!File.Exists(input))
            return (null, "Không tìm thấy file Excel nguồn.");
        if (!Directory.Exists(Path.GetDirectoryName(output)))
            return (null, "Thư mục lưu kết quả không tồn tại.");
        if (File.Exists(output) && !request.ConfirmOutputOverwrite)
            return (null, "File kết quả đã tồn tại. Cần xác nhận ghi đè trước khi tiếp tục.");
        return (new(input, output, isXlsx), null);
    }
}

/// <summary>Temporary files for one install; everything is removed on dispose (best effort).</summary>
internal sealed class InstallWorkArea : IDisposable
{
    private readonly string _packageDirectory;

    public InstallWorkArea(InstallPaths paths)
    {
        var tempBase = Path.Combine(Path.GetDirectoryName(paths.Output)!, $".tool-vba-{Guid.NewGuid():N}.tmp");
        TempPath = tempBase + (paths.IsXlsx ? ".xlsx" : ".xlsm");
        SavedPath = paths.IsXlsx ? tempBase + ".xlsm" : TempPath;
        ConvertsToMacroEnabled = paths.IsXlsx;
        _packageDirectory = Path.Combine(Path.GetTempPath(), $"tool-vba-pkg-{Guid.NewGuid():N}");
    }

    public string TempPath { get; }
    public string SavedPath { get; }
    public bool ConvertsToMacroEnabled { get; }

    /// <summary>
    /// Copies the package files into a private folder and re-verifies them, so a cache refresh or a
    /// change on disk cannot alter what Excel imports.
    /// </summary>
    public VbaPackage? StagePackage(VbaPackage package, VbaPackageReader reader)
    {
        Directory.CreateDirectory(_packageDirectory);
        string Copy(string source)
        {
            var target = Path.Combine(_packageDirectory, Path.GetFileName(source));
            File.Copy(source, target, overwrite: false);
            return target;
        }

        var components = package.Components.Select(c => c with
        {
            FilePath = Copy(c.FilePath),
            ExtraFilePath = c.ExtraFilePath is null ? null : Copy(c.ExtraFilePath)
        }).ToList();
        var staged = package with { Directory = _packageDirectory, Components = components };
        return reader.Verify(staged).IsValid ? staged : null;
    }

    /// <summary>Each item is removed independently, so one locked file does not keep the others behind.</summary>
    public void Dispose()
    {
        TryDelete(() => { if (File.Exists(TempPath)) File.Delete(TempPath); });
        TryDelete(() => { if (ConvertsToMacroEnabled && File.Exists(SavedPath)) File.Delete(SavedPath); });
        TryDelete(() => { if (Directory.Exists(_packageDirectory)) Directory.Delete(_packageDirectory, recursive: true); });
    }

    private static void TryDelete(Action delete)
    {
        try { delete(); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
