using System.Text.Json;
using System.Text.RegularExpressions;

namespace ToolVba.Core;

/// <summary>
/// The store product and version a workbook is packaged for (DEC-115). Values come from the
/// packaging code copied on the Admin web, or are typed by hand; both go through the same checks,
/// which mirror the store's own rules for <c>product_key</c> and version numbers.
/// </summary>
public sealed record LicenseProduct(string ProductKey, string Version)
{
    /// <summary>The store sets no length limit; this keeps the generated VBA line far below Excel's 1023 characters.</summary>
    public const int MaxProductKeyLength = 200;

    /// <summary>The store sets no limit either; real versions are a few characters, and the VBA line must stay short.</summary>
    public const int MaxVersionLength = 50;

    private const int MaxPackagingCodeLength = 1024;

    /// <summary>The VBE password dialog accepts up to this many characters; the server keeps well under it.</summary>
    public const int MaxVbaPasswordLength = 255;

    // Same patterns as CreateProductDto / CreateVersionDto of the API. [0-9] rather than \d, which in .NET
    // also matches non-ASCII digits; \z rather than $, which would accept a trailing newline.
    private static readonly Regex KeyPattern = new(@"^[a-z0-9]+(?:-[a-z0-9]+)*\z", RegexOptions.CultureInvariant);
    private static readonly Regex VersionPattern = new(@"^[0-9]+\.[0-9]+\.[0-9]+\z", RegexOptions.CultureInvariant);

    private const string InvalidKey =
        "Mã sản phẩm (productKey) chỉ gồm chữ thường không dấu, số và dấu '-' giữa các phần, ví dụ test-01.";
    private const string InvalidVersion = "Phiên bản phải có dạng x.y.z, ví dụ 1.0.0.";
    private const string InvalidPackagingCode =
        "Mã đóng gói không đúng. Trên Admin web, mở sản phẩm, bấm \"Sao chép mã đóng gói\" rồi dán lại vào đây.";
    private const string InvalidVbaPassword =
        "Mã đóng gói có trường vbaPassword không hợp lệ: cần chuỗi ASCII in được, 1–255 ký tự.";

    /// <summary>Checks typed values. The key is trimmed and lowercased, as the store does when a product is created.</summary>
    public static (LicenseProduct? Product, string? Error) FromValues(string? productKey, string? version)
    {
        var key = productKey?.Trim().ToLowerInvariant() ?? "";
        if (key.Length is 0 or > MaxProductKeyLength || !KeyPattern.IsMatch(key)) return (null, InvalidKey);
        var number = version?.Trim() ?? "";
        if (number.Length > MaxVersionLength || !VersionPattern.IsMatch(number)) return (null, InvalidVersion);
        return (new LicenseProduct(key, number), null);
    }

    /// <summary>
    /// Reads exactly <c>{"productKey":"…","version":"…"}</c>: an object with those two string fields,
    /// each once, and nothing else.
    /// </summary>
    /// <summary>The product in a packaging code, ignoring any VBA lock password; see <see cref="ReadPackagingCode"/>.</summary>
    public static (LicenseProduct? Product, string? Error) FromPackagingCode(string? text)
    {
        var (code, error) = ReadPackagingCode(text);
        return (code?.Product, error);
    }

    /// <summary>
    /// Reads <c>{"productKey":"…","version":"…"}</c> with an optional <c>"vbaPassword":"…"</c>: an object with
    /// those fields, each at most once, and nothing else. The password, when present, is the server-generated
    /// VBA lock password for this product and version (DEC-116); absent means the file is not locked.
    /// </summary>
    public static (PackagingCode? Code, string? Error) ReadPackagingCode(string? text)
    {
        // Some editors put a byte order mark on the clipboard; it is not whitespace for Trim().
        var trimmed = text?.Trim().TrimStart('﻿').Trim() ?? "";
        if (trimmed.Length is 0 or > MaxPackagingCodeLength) return (null, InvalidPackagingCode);
        try
        {
            using var document = JsonDocument.Parse(trimmed);
            var (ok, key, version, password) = ReadFields(document.RootElement);
            if (!ok || key is null || version is null) return (null, InvalidPackagingCode);
            var (product, productError) = FromValues(key, version);
            if (product is null) return (null, productError);
            if (password is not null && !IsValidVbaPassword(password)) return (null, InvalidVbaPassword);
            return (new PackagingCode(product, password), null);
        }
        // A lone surrogate, raw or escaped, makes the parser or GetString() throw ArgumentException/InvalidOperationException.
        catch (Exception ex) when (ex is JsonException or ArgumentException or InvalidOperationException)
        {
            return (null, InvalidPackagingCode);
        }
    }

    private static bool IsValidVbaPassword(string value) =>
        value.Length is > 0 and <= MaxVbaPasswordLength && value.All(c => c is >= ' ' and <= '~');

    /// <summary>
    /// The recognised fields, with <c>Ok</c> false when the object has any other shape (a non-string value, or
    /// an unknown or repeated field). <c>Password</c> is null when the optional <c>vbaPassword</c> is absent.
    /// </summary>
    private static (bool Ok, string? Key, string? Version, string? Password) ReadFields(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object) return (false, null, null, null);
        string? key = null;
        string? version = null;
        string? password = null;
        foreach (var property in root.EnumerateObject())
        {
            if (property.Value.ValueKind != JsonValueKind.String) return (false, null, null, null);
            switch (property.Name)
            {
                case "productKey" when key is null:
                    key = property.Value.GetString();
                    break;
                case "version" when version is null:
                    version = property.Value.GetString();
                    break;
                case "vbaPassword" when password is null:
                    password = property.Value.GetString();
                    break;
                default:
                    return (false, null, null, null);
            }
        }
        return (true, key, version, password);
    }
}

/// <summary>A parsed packaging code: the store product plus the optional server-generated VBA lock password (DEC-116).</summary>
public sealed record PackagingCode(LicenseProduct Product, string? VbaPassword);
