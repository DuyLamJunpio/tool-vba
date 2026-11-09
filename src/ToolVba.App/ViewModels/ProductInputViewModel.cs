using ToolVba.Core;

namespace ToolVba.App.ViewModels;

/// <summary>
/// The store product a file is packaged for (DEC-115): paste the packaging code copied on the Admin web,
/// which fills both fields, or type the product key and version by hand. Both go through LicenseProduct.
/// </summary>
internal sealed class ProductInputViewModel : ObservableObject
{
    private string _packagingCode = "";
    private string _productKey = "";
    private string _version = "";
    private string? _codeError;
    private bool _filling;

    public event EventHandler? ProductChanged;

    public string PackagingCodeText
    {
        get => _packagingCode;
        set
        {
            if (!Set(ref _packagingCode, value)) return;
            if (string.IsNullOrWhiteSpace(value))
            {
                CodeError = null;
                return;
            }
            var (product, error) = LicenseProduct.FromPackagingCode(value);
            CodeError = error;
            if (product is not null) Fill(product);
        }
    }

    public string ProductKeyText
    {
        get => _productKey;
        set { if (Set(ref _productKey, value) && !_filling) Changed(); }
    }

    public string VersionText
    {
        get => _version;
        set { if (Set(ref _version, value) && !_filling) Changed(); }
    }

    /// <summary>
    /// The valid product, or null while a field is empty or wrong, or while the pasted code is wrong: the fields
    /// may still hold an earlier code, and packaging that one while the screen shows an error would mislead.
    /// </summary>
    public LicenseProduct? Product => CodeError is null ? LicenseProduct.FromValues(_productKey, _version).Product : null;

    /// <summary>What is wrong with the pasted code or the typed values; null while nothing has been typed.</summary>
    public string? Error =>
        CodeError ?? (_productKey.Length == 0 && _version.Length == 0 ? null : LicenseProduct.FromValues(_productKey, _version).Error);

    private string? CodeError
    {
        get => _codeError;
        set
        {
            if (!Set(ref _codeError, value)) return;
            OnPropertyChanged(nameof(Error));
            OnPropertyChanged(nameof(Product));
            ProductChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private void Fill(LicenseProduct product)
    {
        _filling = true;
        try
        {
            ProductKeyText = product.ProductKey;
            VersionText = product.Version;
        }
        finally
        {
            _filling = false;
        }
        Changed();
    }

    private void Changed()
    {
        OnPropertyChanged(nameof(Product));
        OnPropertyChanged(nameof(Error));
        ProductChanged?.Invoke(this, EventArgs.Empty);
    }
}
