using System.IO.Compression;
using System.Security.Cryptography;
using OpenMcdf;

namespace ToolVba.Core;

/// <summary>
/// Locks the VBA project of a packaged .xlsm with a password (DEC-116), so Excel asks for it before showing
/// the VBA editor — the deterrent against a user opening Alt+F11 and deleting the licensing module. The
/// protection is written straight into <c>xl/vbaProject.bin</c> because Excel's object model cannot set it;
/// the byte encoding lives in <see cref="VbaProjectProtection"/>. Call this only after the workbook has been
/// verified: once the project is locked, its code can no longer be read back for checking.
/// </summary>
public static class VbaProjectLocker
{
    private const string VbaProjectEntry = "xl/vbaProject.bin";
    private const string ProjectStreamName = "PROJECT";

    /// <summary>Locks the VBA project of the .xlsm at <paramref name="workbookPath"/> in place.</summary>
    public static void Lock(string workbookPath, string password)
    {
        ArgumentException.ThrowIfNullOrEmpty(workbookPath);
        ArgumentException.ThrowIfNullOrEmpty(password);

        using var zip = ZipFile.Open(workbookPath, ZipArchiveMode.Update);
        var entry = zip.GetEntry(VbaProjectEntry)
            ?? throw new InvalidDataException("Workbook không có xl/vbaProject.bin để khoá VBA project.");

        byte[] original;
        using (var source = entry.Open())
        using (var buffer = new MemoryStream())
        {
            source.CopyTo(buffer);
            original = buffer.ToArray();
        }

        var locked = LockVbaProject(original, password);
        // A zip entry cannot be resized in place; replace it. Entry order does not matter to Excel.
        entry.Delete();
        var replacement = zip.CreateEntry(VbaProjectEntry);
        using var target = replacement.Open();
        target.Write(locked, 0, locked.Length);
    }

    /// <summary>Returns a copy of the vbaProject.bin bytes with the PROJECT stream locked by the password.</summary>
    public static byte[] LockVbaProject(byte[] vbaProjectBin, string password)
    {
        using var storageStream = new MemoryStream();
        storageStream.Write(vbaProjectBin, 0, vbaProjectBin.Length);
        storageStream.Position = 0;
        using (var root = RootStorage.Open(storageStream))
        {
            using var project = root.OpenStream(ProjectStreamName);
            var current = new byte[project.Length];
            project.ReadExactly(current, 0, current.Length);
            var locked = VbaProjectProtection.LockProjectStream(current, password, RandomNumberGenerator.GetBytes);
            project.SetLength(locked.Length);
            project.Position = 0;
            project.Write(locked, 0, locked.Length);
        }
        return storageStream.ToArray();
    }
}
