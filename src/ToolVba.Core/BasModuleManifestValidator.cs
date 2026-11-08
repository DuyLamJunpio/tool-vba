using System.Security.Cryptography;
using System.Text.Json;

namespace ToolVba.Core;

public sealed record BasModuleManifest(string ModuleName, string Sha256);

public sealed record BasModuleValidationResult(bool IsValid, string? Error, string? ModuleName, string? Sha256);

public sealed class BasModuleManifestValidator
{
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    public BasModuleValidationResult Validate(string modulePath, string manifestPath)
    {
        if (!Path.GetExtension(modulePath).Equals(".bas", StringComparison.OrdinalIgnoreCase))
            return Invalid("Module phải là file .bas.");
        if (!Path.GetExtension(manifestPath).Equals(".json", StringComparison.OrdinalIgnoreCase))
            return Invalid("Manifest module phải là file .json.");
        if (!File.Exists(modulePath))
            return Invalid("Không tìm thấy file module .bas.");
        if (!File.Exists(manifestPath))
            return Invalid("Không tìm thấy manifest phê duyệt module.");

        try
        {
            using var manifestStream = File.OpenRead(manifestPath);
            var manifest = JsonSerializer.Deserialize<BasModuleManifest>(manifestStream, JsonOptions);
            if (manifest is null || string.IsNullOrWhiteSpace(manifest.ModuleName) || manifest.ModuleName.Length > VbaSourceRules.MaxComponentNameLength || string.IsNullOrWhiteSpace(manifest.Sha256))
                return Invalid("Manifest thiếu tên module hợp lệ.");
            if (manifest.Sha256.Length != 64 || !manifest.Sha256.All(Uri.IsHexDigit))
                return Invalid("SHA-256 trong manifest không hợp lệ.");

            using var moduleStream = File.OpenRead(modulePath);
            if (moduleStream.Length == 0)
                return Invalid("Module .bas đang trống.");
            var actualHash = SHA256.HashData(moduleStream);
            var expectedHash = Convert.FromHexString(manifest.Sha256);
            if (!CryptographicOperations.FixedTimeEquals(actualHash, expectedHash))
                return Invalid("SHA-256 của module không khớp manifest phê duyệt.");

            if (VbaSourceRules.FindReservedInputDeclaration(File.ReadLines(modulePath)) is not null)
                return Invalid("Module dùng từ khóa VBA 'Input' làm tên tham số hoặc biến. Hãy đổi tên rồi cập nhật manifest.");

            return new(true, null, manifest.ModuleName, Convert.ToHexString(actualHash));
        }
        catch (JsonException)
        {
            return Invalid("Manifest không phải JSON hợp lệ.");
        }
        catch (FormatException)
        {
            return Invalid("SHA-256 trong manifest không hợp lệ.");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return Invalid("Không thể đọc module hoặc manifest. Kiểm tra đường dẫn và quyền truy cập.");
        }
    }

    private static BasModuleValidationResult Invalid(string message) => new(false, message, null, null);
}
