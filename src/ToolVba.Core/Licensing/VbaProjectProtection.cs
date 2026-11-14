using System.Security.Cryptography;
using System.Text;

namespace ToolVba.Core;

/// <summary>
/// The MS-OVBA project-lock encoding (sections 2.3.1, 2.4.3, 2.4.4): turns a password into the
/// <c>CMG</c>/<c>DPB</c>/<c>GC</c> values of the PROJECT stream so Excel asks for the password before it
/// shows the VBA project (DEC-116). This is pure byte logic, isolated from file I/O so it can be
/// unit-tested against the spec's own worked example. The file I/O lives in the VBA project locker.
/// </summary>
public static class VbaProjectProtection
{
    private const byte DataEncryptionVersion = 2; // MS-OVBA 2.4.3.2: MUST be 2.

    // A project whose DPB stores a password hash MUST carry the null CLSID (MS-OVBA 2.3.1.16).
    private const string NullProjectId = "{00000000-0000-0000-0000-000000000000}";

    // fVBEProtected (bit 2); the value Excel itself writes for "Lock project for viewing".
    private static readonly byte[] ProtectedState = BitConverter.GetBytes(0x00000004u);
    private static readonly byte[] NotVisible = [0x00];

    /// <summary>
    /// Rewrites the PROJECT stream bytes so the project is password-protected and locked for viewing.
    /// The stream is treated as Latin1, which round-trips every byte, so non-ASCII content on other
    /// lines is preserved exactly; only the ID/CMG/DPB/GC lines (all ASCII) are replaced.
    /// </summary>
    /// <param name="random">Returns n random bytes; a test passes a deterministic source.</param>
    public static byte[] LockProjectStream(byte[] projectStream, string password, Func<int, byte[]> random)
    {
        ArgumentNullException.ThrowIfNull(projectStream);
        ArgumentNullException.ThrowIfNull(password);
        ArgumentNullException.ThrowIfNull(random);

        var salt = random(4);
        var hash = BuildPasswordHashStructure(password, salt);
        var cmg = Convert.ToHexString(Encrypt(ProtectedState, NullProjectId, Seed(random)));
        var dpb = Convert.ToHexString(Encrypt(hash, NullProjectId, Seed(random)));
        var gc = Convert.ToHexString(Encrypt(NotVisible, NullProjectId, Seed(random)));

        var lines = Encoding.Latin1.GetString(projectStream).Split("\r\n").ToList();
        SetProperty(lines, "ID", $"\"{NullProjectId}\"");
        SetProperty(lines, "CMG", $"\"{cmg}\"");
        SetProperty(lines, "DPB", $"\"{dpb}\"");
        SetProperty(lines, "GC", $"\"{gc}\"");
        return Encoding.Latin1.GetBytes(string.Join("\r\n", lines));
    }

    private static byte Seed(Func<int, byte[]> random) => random(1)[0];

    /// <summary>Replaces the <c>KEY=</c> line, or inserts it just after the ID line when absent.</summary>
    private static void SetProperty(List<string> lines, string key, string value)
    {
        var prefix = key + "=";
        for (var i = 0; i < lines.Count; i++)
        {
            if (!lines[i].StartsWith(prefix, StringComparison.Ordinal)) continue;
            lines[i] = prefix + value;
            return;
        }
        var idIndex = lines.FindIndex(l => l.StartsWith("ID=", StringComparison.Ordinal));
        lines.Insert(idIndex >= 0 ? idIndex + 1 : 0, prefix + value);
    }

    /// <summary>MS-OVBA 2.4.3.2 Data Encryption. <paramref name="seed"/> also fills the ignored bytes.</summary>
    public static byte[] Encrypt(byte[] data, string projectId, byte seed)
    {
        var projKey = ProjKey(projectId);
        var versionEnc = (byte)(DataEncryptionVersion ^ seed);
        var projKeyEnc = (byte)(projKey ^ seed);
        var output = new List<byte>(7 + data.Length) { seed, versionEnc, projKeyEnc };

        byte encryptedByte1 = projKeyEnc, encryptedByte2 = versionEnc, unencryptedByte1 = projKey;
        void EncryptByte(byte b)
        {
            var e = (byte)(b ^ (byte)(encryptedByte2 + unencryptedByte1));
            output.Add(e);
            encryptedByte2 = encryptedByte1;
            encryptedByte1 = e;
            unencryptedByte1 = b;
        }

        var ignoredLength = (seed & 6) / 2;
        for (var i = 0; i < ignoredLength; i++) EncryptByte(seed);
        foreach (var b in BitConverter.GetBytes((uint)data.Length)) EncryptByte(b);
        foreach (var b in data) EncryptByte(b);
        return [.. output];
    }

    /// <summary>MS-OVBA 2.4.3.3 Data Decryption. Returns the Data field; used by tests to assert round-trips.</summary>
    public static byte[] Decrypt(byte[] encrypted)
    {
        var seed = encrypted[0];
        var decrypted = new byte[encrypted.Length - 1];
        decrypted[0] = (byte)(encrypted[1] ^ seed); // version
        decrypted[1] = (byte)(encrypted[2] ^ seed); // projKey
        for (var i = 2; i < encrypted.Length - 1; i++)
            decrypted[i] = (byte)(encrypted[i + 1] ^ (byte)(encrypted[i - 1] + decrypted[i - 1]));

        var ignoredLength = (seed & 6) / 2;
        var dataLength = BitConverter.ToInt32(decrypted, ignoredLength + 2);
        var data = new byte[dataLength];
        Array.Copy(decrypted, 6 + ignoredLength, data, 0, dataLength);
        return data;
    }

    /// <summary>MS-OVBA 2.4.3.2: the 1-byte project key is the sum (mod 256) of the ID string, braces included.</summary>
    public static byte ProjKey(string projectId)
    {
        byte key = 0;
        foreach (var c in projectId) key += (byte)c;
        return key;
    }

    /// <summary>
    /// The 29-byte Password Hash Data Structure (MS-OVBA 2.4.4.1): 0xFF, three grbit bytes, the null-encoded
    /// 4-byte salt, the null-encoded 20-byte SHA-1 of (password bytes + salt), and a 0x00 terminator. The
    /// password is validated as printable ASCII upstream, so ASCII bytes match the MBCS the spec expects.
    /// </summary>
    public static byte[] BuildPasswordHashStructure(string password, byte[] salt)
    {
        if (salt.Length != 4) throw new ArgumentException("Salt must be 4 bytes.", nameof(salt));
        var passwordBytes = Encoding.ASCII.GetBytes(password);
        var toHash = new byte[passwordBytes.Length + 4];
        Buffer.BlockCopy(passwordBytes, 0, toHash, 0, passwordBytes.Length);
        Buffer.BlockCopy(salt, 0, toHash, passwordBytes.Length, 4);
        var hash = SHA1.HashData(toHash); // 20 bytes

        var grbit = new byte[3];
        var keyNoNulls = (byte[])salt.Clone();
        var hashNoNulls = (byte[])hash.Clone();
        for (var i = 0; i < 24; i++)
        {
            var bit = (byte)(128 >> (i % 8)); // MSB first within each grbit byte
            if (i < 4)
            {
                if (keyNoNulls[i] == 0) keyNoNulls[i] = 1; else grbit[0] |= bit;
            }
            else
            {
                if (hashNoNulls[i - 4] == 0) hashNoNulls[i - 4] = 1; else grbit[i / 8] |= bit;
            }
        }

        var structure = new byte[29];
        structure[0] = 0xFF;
        structure[1] = grbit[0];
        structure[2] = grbit[1];
        structure[3] = grbit[2];
        Buffer.BlockCopy(keyNoNulls, 0, structure, 4, 4);
        Buffer.BlockCopy(hashNoNulls, 0, structure, 8, 20);
        structure[28] = 0x00;
        return structure;
    }
}
