using System.Security.Cryptography;
using System.Text;

namespace ToolVba.Core;

/// <summary>
/// Verifies ECDSA P-256 / SHA-256 release signatures in IEEE P1363 format
/// (docs/tool-release-api-contract.md §1.2). Only the public key is embedded in the app.
/// </summary>
public sealed class ReleaseSignatureVerifier
{
    private const int P256SignatureBytes = 64;
    private const string NistP256Oid = "1.2.840.10045.3.1.7";
    private readonly byte[] _subjectPublicKeyInfo;

    public ReleaseSignatureVerifier(string publicKeyPem)
    {
        using var key = ECDsa.Create();
        try
        {
            key.ImportFromPem(publicKeyPem);
        }
        catch (CryptographicException ex)
        {
            throw new ArgumentException("Public key phát hành không hợp lệ.", nameof(publicKeyPem), ex);
        }
        if (key.KeySize != 256 || key.ExportParameters(false).Curve.Oid?.Value != NistP256Oid)
            throw new ArgumentException("Public key phát hành phải là ECDSA P-256.", nameof(publicKeyPem));
        _subjectPublicKeyInfo = key.ExportSubjectPublicKeyInfo();
    }

    public static string VbaPackagePayload(string version, string sha256, string? minAppVersion) =>
        $"VBA_PACKAGE|{version}|{sha256}|{minAppVersion ?? ""}";

    public bool Verify(string payload, string? signatureBase64)
    {
        if (string.IsNullOrWhiteSpace(signatureBase64)) return false;
        var signature = new byte[P256SignatureBytes];
        if (!Convert.TryFromBase64String(signatureBase64, signature, out var written) || written != P256SignatureBytes)
            return false;

        using var key = ECDsa.Create();
        key.ImportSubjectPublicKeyInfo(_subjectPublicKeyInfo, out _);
        return key.VerifyData(Encoding.UTF8.GetBytes(payload), signature, HashAlgorithmName.SHA256,
            DSASignatureFormat.IeeeP1363FixedFieldConcatenation);
    }
}
