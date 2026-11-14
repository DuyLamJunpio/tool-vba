using System.Text;
using ToolVba.Core;

namespace ToolVba.Tests;

/// <summary>
/// Validates the MS-OVBA encoding against the spec's own PROJECT Stream Example (section 3.1.6): decryption
/// is deterministic, so the example's CMG/DPB/GC must decode to the documented Data, and CMG must re-encrypt
/// byte-for-byte. Encryption of DPB/GC is non-deterministic (random seed and salt), so those are checked by
/// round-trip instead.
/// </summary>
public sealed class VbaProjectProtectionTests
{
    private const string ExampleId = "{917DED54-440B-4FD1-A5C1-74ACF261E600}";
    private const string NullId = "{00000000-0000-0000-0000-000000000000}";

    [Fact]
    public void ProjKey_sums_the_id_string_including_braces()
    {
        Assert.Equal(0xDF, VbaProjectProtection.ProjKey(ExampleId));
        Assert.Equal(0xAC, VbaProjectProtection.ProjKey(NullId));
    }

    [Theory]
    [InlineData("0705D8E3D8EDDBF1DBF1DBF1DBF1", new byte[] { 0, 0, 0, 0 })]
    [InlineData("0E0CD1ECDFF4E7F5E7F5E7", new byte[] { 0 })]
    [InlineData("1517CAF1D6F9D7F9D706", new byte[] { 0xFF })]
    public void Decrypts_the_spec_example_blobs(string hex, byte[] expected)
    {
        Assert.Equal(expected, VbaProjectProtection.Decrypt(Convert.FromHexString(hex)));
    }

    [Fact]
    public void Encrypts_the_spec_cmg_example_byte_for_byte()
    {
        var encrypted = VbaProjectProtection.Encrypt([0, 0, 0, 0], ExampleId, 0x07);

        Assert.Equal("0705D8E3D8EDDBF1DBF1DBF1DBF1", Convert.ToHexString(encrypted));
    }

    [Theory]
    [InlineData((byte)0x00)]
    [InlineData((byte)0x07)]
    [InlineData((byte)0x0E)]
    [InlineData((byte)0x15)]
    [InlineData((byte)0xFF)]
    public void Encrypt_then_decrypt_round_trips_any_seed(byte seed)
    {
        byte[] data = [1, 2, 3, 4, 5, 0, 250, 0, 0];

        var roundTripped = VbaProjectProtection.Decrypt(VbaProjectProtection.Encrypt(data, ExampleId, seed));

        Assert.Equal(data, roundTripped);
    }

    [Fact]
    public void Password_hash_structure_has_the_spec_shape_and_encodes_nulls()
    {
        var structure = VbaProjectProtection.BuildPasswordHashStructure("test", [0, 0, 0, 0]);

        Assert.Equal(29, structure.Length);
        Assert.Equal(0xFF, structure[0]);
        Assert.Equal(0x00, structure[28]);
        // Every salt byte was 0x00, so KeyNoNulls is all 0x01 and the four key grbit bits are clear.
        Assert.Equal(new byte[] { 1, 1, 1, 1 }, structure[4..8]);
        Assert.Equal(0, structure[1] & 0xF0);
        // No encoded byte may be a literal 0x00 (that is the whole point of null-encoding).
        Assert.DoesNotContain((byte)0, structure[4..28]);
    }

    [Fact]
    public void Lock_project_stream_replaces_the_four_lines_and_keeps_the_rest()
    {
        var project = Encoding.Latin1.GetBytes(
            "ID=\"{5DD90D76-4904-47A2-AF0D-D69B4673604E}\"\r\n" +
            "Document=ThisWorkbook/&H00000000\r\n" +
            "CMG=\"AAAA\"\r\n" +
            "DPB=\"BBBB\"\r\n" +
            "GC=\"CCCC\"\r\n" +
            "[Host Extender Info]\r\n");

        var locked = VbaProjectProtection.LockProjectStream(project, "QWxhZGRpbg==", Deterministic());
        var text = Encoding.Latin1.GetString(locked);

        Assert.Contains($"ID=\"{NullId}\"", text, StringComparison.Ordinal);
        // Lines that are not ID/CMG/DPB/GC are untouched.
        Assert.Contains("Document=ThisWorkbook/&H00000000\r\n", text, StringComparison.Ordinal);
        Assert.Contains("[Host Extender Info]", text, StringComparison.Ordinal);
        // The protection blobs decode back to the documented Data.
        Assert.Equal(new byte[] { 4, 0, 0, 0 }, VbaProjectProtection.Decrypt(Convert.FromHexString(Quoted(text, "CMG="))));
        Assert.Equal(new byte[] { 0 }, VbaProjectProtection.Decrypt(Convert.FromHexString(Quoted(text, "GC="))));
        var dpb = VbaProjectProtection.Decrypt(Convert.FromHexString(Quoted(text, "DPB=")));
        Assert.Equal(29, dpb.Length);
        Assert.Equal(0xFF, dpb[0]);
    }

    /// <summary>A predictable byte source so the locked stream is reproducible in the test.</summary>
    private static Func<int, byte[]> Deterministic()
    {
        byte next = 0;
        return count =>
        {
            var bytes = new byte[count];
            for (var i = 0; i < count; i++) bytes[i] = next++;
            return bytes;
        };
    }

    private static string Quoted(string text, string key)
    {
        var start = text.IndexOf(key + "\"", StringComparison.Ordinal) + key.Length + 1;
        var end = text.IndexOf('"', start);
        return text[start..end];
    }
}
