using System.Text.Json;
using System.Text.Json.Serialization;

namespace ToolVba.Core;

/// <summary>Kind of a VBA component inside a package or a workbook project.</summary>
public enum VbaComponentKind
{
    Standard,
    Class,
    Form,
    Workbook,
    Sheet,
    Other
}

/// <summary>One entry of <c>package.json</c> (see docs/tool-release-api-contract.md §1.1).</summary>
public sealed record VbaPackageManifestComponent(
    string Name,
    VbaComponentKind Kind,
    string File,
    string Sha256,
    string? ExtraFile = null,
    string? ExtraSha256 = null,
    string? SheetName = null);

public sealed record VbaPackageManifest(int FormatVersion, IReadOnlyList<VbaPackageManifestComponent> Components)
{
    public const int CurrentFormatVersion = 1;
    public const string FileName = "package.json";

    // Strict like the server validator: exact (camelCase) names and no unknown fields, so the client never
    // reads a component list the server did not validate.
    internal static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web)
    {
        PropertyNameCaseInsensitive = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        Converters = { new JsonStringEnumConverter<VbaComponentKind>(JsonNamingPolicy.CamelCase, allowIntegerValues: false) },
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = true
    };
}

/// <summary>A verified component file extracted on disk.</summary>
public sealed record VbaPackageComponentFile(
    string Name,
    VbaComponentKind Kind,
    string FilePath,
    string Sha256,
    string? ExtraFilePath = null,
    string? ExtraSha256 = null,
    string? SheetName = null)
{
    public bool IsCodeModule => Kind is VbaComponentKind.Standard or VbaComponentKind.Class or VbaComponentKind.Form;
    public bool IsDocumentModule => Kind is VbaComponentKind.Workbook or VbaComponentKind.Sheet;
}

/// <summary>A verified, extracted company VBA package.</summary>
public sealed record VbaPackage(string Version, string Directory, IReadOnlyList<VbaPackageComponentFile> Components)
{
    public IEnumerable<VbaPackageComponentFile> CodeModules => Components.Where(c => c.IsCodeModule);
    public IEnumerable<VbaPackageComponentFile> DocumentModules => Components.Where(c => c.IsDocumentModule);
}

public sealed record VbaPackageReadResult(bool IsValid, string? Error, VbaPackage? Package)
{
    public static VbaPackageReadResult Valid(VbaPackage package) => new(true, null, package);
    public static VbaPackageReadResult Invalid(string error) => new(false, error, null);
}

/// <summary>A staged component file used to build a package.</summary>
/// <param name="SheetName">Tab name of a sheet component, used when the target workbook's CodeName differs.</param>
public sealed record VbaPackageSource(string Name, VbaComponentKind Kind, string FilePath, string? ExtraFilePath = null, string? SheetName = null);

public sealed record VbaPackageBuildResult(bool IsValid, string? Error, IReadOnlyList<VbaPackageManifestComponent> Components)
{
    public static VbaPackageBuildResult Invalid(string error) => new(false, error, []);
}
