using System.Text;
using OpenMcdf;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class VbaProjectLockerTests
{
    [Fact]
    public void Locks_the_project_stream_inside_a_compound_file()
    {
        var bin = CreateVbaProjectBin(
            "ID=\"{5DD90D76-4904-47A2-AF0D-D69B4673604E}\"\r\n" +
            "Document=ThisWorkbook/&H00000000\r\n" +
            "CMG=\"AAAA\"\r\n" +
            "DPB=\"BBBB\"\r\n" +
            "GC=\"CCCC\"\r\n");

        var locked = VbaProjectLocker.LockVbaProject(bin, "QWxhZGRpbg==");

        var text = Encoding.Latin1.GetString(ReadProjectStream(locked));
        Assert.Contains("ID=\"{00000000-0000-0000-0000-000000000000}\"", text, StringComparison.Ordinal);
        Assert.Contains("Document=ThisWorkbook/&H00000000", text, StringComparison.Ordinal);
        Assert.Equal(new byte[] { 4, 0, 0, 0 }, VbaProjectProtection.Decrypt(Convert.FromHexString(Quoted(text, "CMG="))));
        Assert.Equal(new byte[] { 0 }, VbaProjectProtection.Decrypt(Convert.FromHexString(Quoted(text, "GC="))));
        Assert.Equal(29, VbaProjectProtection.Decrypt(Convert.FromHexString(Quoted(text, "DPB="))).Length);
    }

    private static byte[] CreateVbaProjectBin(string projectText)
    {
        using var ms = new MemoryStream();
        using (var root = RootStorage.Create(ms))
        {
            using var project = root.CreateStream("PROJECT");
            var bytes = Encoding.Latin1.GetBytes(projectText);
            project.Write(bytes, 0, bytes.Length);
        }
        return ms.ToArray();
    }

    private static byte[] ReadProjectStream(byte[] bin)
    {
        using var ms = new MemoryStream();
        ms.Write(bin, 0, bin.Length);
        ms.Position = 0;
        using var root = RootStorage.Open(ms);
        using var project = root.OpenStream("PROJECT");
        var bytes = new byte[project.Length];
        project.ReadExactly(bytes, 0, bytes.Length);
        return bytes;
    }

    private static string Quoted(string text, string key)
    {
        var start = text.IndexOf(key + "\"", StringComparison.Ordinal) + key.Length + 1;
        var end = text.IndexOf('"', start);
        return text[start..end];
    }
}
