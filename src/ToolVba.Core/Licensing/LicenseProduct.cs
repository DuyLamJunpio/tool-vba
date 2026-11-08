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

    // Same patterns as CreateProductDto / CreateVersionDto of the API. [0-9] rather than \d, which in .NET
    // also matches non-ASCII digits; \z rather than $, which would accept a trailing newline.
    private static readonly Regex KeyPattern = new(@"^[a-z0-9]+(?:-[a-z0-9]+)*\z", RegexOptions.CultureInvariant);
    private static readonly Regex VersionPattern = new(@"^[0-9]+\.[0-9]+\.[0-9]+\z", RegexOptions.CultureInvariant);

    private const string InvalidKey =
        "Mã sản phẩm (productKey) chỉ gồm chữ thường không dấu, số và dấu '-' giữa các phần, ví dụ test-01.";
    private const string InvalidVersion = "Phiên bản phải có dạng x.y.z, ví dụ 1.0.0.";
    private const string InvalidPackagingCode =
        "Mã đóng gói không đúng. Trên Admin web, mở sản phẩm, bấm \"Sao chép mã đóng gói\" rồi dán lại vào đây.";

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
    public static (LicenseProduct? Product, string? Error) FromPackagingCode(string? text)
    {
        // Some editors put a byte order mark on the clipboard; it is not whitespace for Trim().
        var trimmed = text?.Trim().TrimStart('﻿').Trim() ?? "";
        if (trimmed.Length is 0 or > MaxPackagingCodeLength) return (null, InvalidPackagingCode);
        try
        {
            using var document = JsonDocument.Parse(trimmed);
            var (key, version) = ReadFields(document.RootElement);
            return key is null || version is null ? (null, InvalidPackagingCode) : FromValues(key, version);
        }
        // A lone surrogate, raw or escaped, makes the parser or GetString() throw ArgumentException/InvalidOperationException.
        catch (Exception ex) when (ex is JsonException or ArgumentException or InvalidOperationException)
        {
            return (null, InvalidPackagingCode);
        }
    }

    /// <summary>Both fields, or nulls when the object has any other shape.</summary>
    private static (string? Key, string? Version) ReadFields(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object) return (null, null);
        string? key = null;
        string? version = null;
        foreach (var property in root.EnumerateObject())
        {
            if (property.Value.ValueKind != JsonValueKind.String) return (null, null);
            switch (property.Name)
            {
                case "productKey" when key is null:
                    key = property.Value.GetString();
                    break;
                case "version" when version is null:
                    version = property.Value.GetString();
                    break;
                default:
                    return (null, null);
            }
        }
        return (key, version);
    }
}
