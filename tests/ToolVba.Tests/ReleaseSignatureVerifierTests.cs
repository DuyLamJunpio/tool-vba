using System.Security.Cryptography;
using System.Text;
using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class ReleaseSignatureVerifierTests
{
    [Fact]
    public void Accepts_p1363_signature_made_with_matching_private_key()
    {
        using var key = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        var payload = ReleaseSignatureVerifier.VbaPackagePayload("2026.11.7.1", new string('a', 64), "1.2.0");
        var signature = Sign(key, payload);

        var verifier = new ReleaseSignatureVerifier(key.ExportSubjectPublicKeyInfoPem());

        Assert.Equal("VBA_PACKAGE|2026.11.7.1|" + new string('a', 64) + "|1.2.0", payload);
        Assert.True(verifier.Verify(payload, signature));
    }

    [Fact]
    public void Rejects_changed_payload_other_key_and_malformed_signature()
    {
        using var key = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        using var other = ECDsa.Create(ECCurve.NamedCurves.nistP256);
        var payload = ReleaseSignatureVerifier.VbaPackagePayload("2026.11.7.1", new string('b', 64), null);
        var signature = Sign(key, payload);
        var verifier = new ReleaseSignatureVerifier(key.ExportSubjectPublicKeyInfoPem());

        Assert.EndsWith("|", payload);
        Assert.False(verifier.Verify(payload.Replace("2026.11.7.1", "2026.11.7.2"), signature));
        Assert.False(new ReleaseSignatureVerifier(other.ExportSubjectPublicKeyInfoPem()).Verify(payload, signature));
        Assert.False(verifier.Verify(payload, "not base64!"));
        Assert.False(verifier.Verify(payload, Convert.ToBase64String(new byte[10])));
    }

    [Fact]
    public void Refuses_invalid_public_key_at_construction()
    {
        Assert.ThrowsAny<ArgumentException>(() => new ReleaseSignatureVerifier("not a key"));
    }

    internal static string Sign(ECDsa key, string payload) =>
        Convert.ToBase64String(key.SignData(Encoding.UTF8.GetBytes(payload), HashAlgorithmName.SHA256,
            DSASignatureFormat.IeeeP1363FixedFieldConcatenation));
}
