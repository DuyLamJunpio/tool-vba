using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class LicenseProductTests
{
    [Fact]
    public void Accepts_the_packaging_code_copied_from_the_admin()
    {
        var (product, error) = LicenseProduct.FromPackagingCode("{\"productKey\":\"test-01\",\"version\":\"1.0.0\"}");

        Assert.Null(error);
        Assert.Equal(new LicenseProduct("test-01", "1.0.0"), product);
    }

    [Fact]
    public void Ignores_whitespace_and_line_breaks_around_a_pasted_code()
    {
        var (product, error) = LicenseProduct.FromPackagingCode("  \r\n{ \"version\": \"2.10.3\", \"productKey\": \"bao-cao-thue\" }\r\n ");

        Assert.Null(error);
        Assert.Equal(new LicenseProduct("bao-cao-thue", "2.10.3"), product);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("test-01")]
    [InlineData("[\"test-01\",\"1.0.0\"]")]
    [InlineData("{\"productKey\":\"test-01\"}")]
    [InlineData("{\"version\":\"1.0.0\"}")]
    [InlineData("{\"productKey\":\"test-01\",\"version\":\"1.0.0\",\"id\":\"x\"}")]
    [InlineData("{\"ProductKey\":\"test-01\",\"version\":\"1.0.0\"}")]
    [InlineData("{\"productKey\":1,\"version\":\"1.0.0\"}")]
    [InlineData("{\"productKey\":\"test-01\",\"version\":null}")]
    [InlineData("{\"productKey\":\"test-01\",\"productKey\":\"other\",\"version\":\"1.0.0\"}")]
    [InlineData("{\"productKey\":\"test-01\",\"version\":\"1.0.0\"")]
    [InlineData("{}")]
    [InlineData("{\"productKey\":\"test-01\",\"version\":\"1.0.0\"} extra")]
    [InlineData("{\"productKey\":\"a\\ud800\",\"version\":\"1.0.0\"}")]
    [InlineData("{\"productKey\":\"a\ud800\",\"version\":\"1.0.0\"}")]
    [InlineData("\ud800")]
    public void Refuses_a_packaging_code_that_is_not_exactly_key_and_version(string text)
    {
        var (product, error) = LicenseProduct.FromPackagingCode(text);

        Assert.Null(product);
        Assert.False(string.IsNullOrWhiteSpace(error));
    }

    [Fact]
    public void Refuses_a_packaging_code_longer_than_any_real_one()
    {
        var text = "{\"productKey\":\"" + new string('a', 2000) + "\",\"version\":\"1.0.0\"}";

        var (product, error) = LicenseProduct.FromPackagingCode(text);

        Assert.Null(product);
        Assert.NotNull(error);
    }

    [Fact]
    public void Checks_the_values_inside_a_packaging_code_like_typed_values()
    {
        var (product, error) = LicenseProduct.FromPackagingCode("{\"productKey\":\"test 01\",\"version\":\"1.0.0\"}");

        Assert.Null(product);
        Assert.Contains("Mã sản phẩm", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("test-01", "1.0.0")]
    [InlineData("a", "0.0.1")]
    [InlineData("excel-cashflow-pro", "10.20.30")]
    public void Accepts_typed_values_in_the_store_formats(string key, string version)
    {
        var (product, error) = LicenseProduct.FromValues(key, version);

        Assert.Null(error);
        Assert.Equal(new LicenseProduct(key, version), product);
    }

    [Fact]
    public void Trims_and_lowercases_a_typed_key_as_the_store_does()
    {
        var (product, error) = LicenseProduct.FromValues("  Test-01 ", " 1.0.0 ");

        Assert.Null(error);
        Assert.Equal(new LicenseProduct("test-01", "1.0.0"), product);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("test_01")]
    [InlineData("test--01")]
    [InlineData("-test")]
    [InlineData("test-")]
    [InlineData("tést")]
    [InlineData("test 01")]
    [InlineData("test\"01")]
    public void Refuses_a_key_outside_the_store_pattern(string? key)
    {
        var (product, error) = LicenseProduct.FromValues(key, "1.0.0");

        Assert.Null(product);
        Assert.Contains("Mã sản phẩm", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Accepts_a_pasted_code_that_starts_with_a_byte_order_mark()
    {
        var (product, error) = LicenseProduct.FromPackagingCode("﻿{\"productKey\":\"test-01\",\"version\":\"1.0.0\"}");

        Assert.Null(error);
        Assert.Equal(new LicenseProduct("test-01", "1.0.0"), product);
    }

    [Fact]
    public void Accepts_the_longest_allowed_key_and_version()
    {
        var key = new string('a', LicenseProduct.MaxProductKeyLength);
        var version = "1.0." + new string('9', LicenseProduct.MaxVersionLength - 4);

        var (product, error) = LicenseProduct.FromValues(key, version);

        Assert.Null(error);
        Assert.Equal(new LicenseProduct(key, version), product);
    }

    [Fact]
    public void Refuses_a_version_too_long_for_one_vba_line()
    {
        var (product, error) = LicenseProduct.FromValues("test-01", "1.0." + new string('9', LicenseProduct.MaxVersionLength - 3));

        Assert.Null(product);
        Assert.Contains("Phiên bản", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Refuses_a_key_too_long_for_one_vba_line()
    {
        var (product, error) = LicenseProduct.FromValues(new string('a', LicenseProduct.MaxProductKeyLength + 1), "1.0.0");

        Assert.Null(product);
        Assert.Contains("Mã sản phẩm", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("1.0")]
    [InlineData("1.0.0.1")]
    [InlineData("v1.0.0")]
    [InlineData("1.0.0-beta")]
    [InlineData("1.a.0")]
    [InlineData("١.٠.٠")]
    public void Refuses_a_version_that_is_not_three_numbers(string? version)
    {
        var (product, error) = LicenseProduct.FromValues("test-01", version);

        Assert.Null(product);
        Assert.Contains("Phiên bản", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Reads_a_packaging_code_with_a_vba_lock_password()
    {
        var (code, error) = LicenseProduct.ReadPackagingCode(
            "{\"productKey\":\"test-01\",\"version\":\"1.0.0\",\"vbaPassword\":\"QWxhZGRpbjpvcGVu\"}");

        Assert.Null(error);
        Assert.Equal(new LicenseProduct("test-01", "1.0.0"), code!.Product);
        Assert.Equal("QWxhZGRpbjpvcGVu", code.VbaPassword);
    }

    [Fact]
    public void Reads_a_packaging_code_without_a_password_as_no_lock()
    {
        var (code, error) = LicenseProduct.ReadPackagingCode("{\"productKey\":\"test-01\",\"version\":\"1.0.0\"}");

        Assert.Null(error);
        Assert.Null(code!.VbaPassword);
    }

    [Fact]
    public void FromPackagingCode_still_returns_the_product_when_a_password_is_present()
    {
        var (product, error) = LicenseProduct.FromPackagingCode(
            "{\"productKey\":\"test-01\",\"version\":\"1.0.0\",\"vbaPassword\":\"QWxhZGRpbg==\"}");

        Assert.Null(error);
        Assert.Equal(new LicenseProduct("test-01", "1.0.0"), product);
    }

    [Theory]
    [InlineData("{\"productKey\":\"test-01\",\"version\":\"1.0.0\",\"vbaPassword\":\"\"}")]        // empty
    [InlineData("{\"productKey\":\"test-01\",\"version\":\"1.0.0\",\"vbaPassword\":\"ab\\tcd\"}")]  // tab is a control char
    [InlineData("{\"productKey\":\"test-01\",\"version\":\"1.0.0\",\"vbaPassword\":\"mật\"}")]      // non-ASCII
    public void Refuses_a_packaging_code_with_an_invalid_vba_password(string text)
    {
        var (code, error) = LicenseProduct.ReadPackagingCode(text);

        Assert.Null(code);
        Assert.Contains("vbaPassword", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Refuses_a_vba_password_longer_than_the_dialog_allows()
    {
        var text = "{\"productKey\":\"test-01\",\"version\":\"1.0.0\",\"vbaPassword\":\"" +
                   new string('a', LicenseProduct.MaxVbaPasswordLength + 1) + "\"}";

        var (code, error) = LicenseProduct.ReadPackagingCode(text);

        Assert.Null(code);
        Assert.Contains("vbaPassword", error, StringComparison.Ordinal);
    }
}
