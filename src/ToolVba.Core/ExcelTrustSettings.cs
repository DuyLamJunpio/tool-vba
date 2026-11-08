using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace ToolVba.Core;

public enum RegistryRoot { CurrentUser, LocalMachine, ClassesRoot }

public interface IRegistryStore
{
    string? ReadString(RegistryRoot root, string subKey, string name);
    int? ReadDword(RegistryRoot root, string subKey, string name);
    void WriteDword(RegistryRoot root, string subKey, string name, int value);
}

public enum VbaTrustState { Enabled, Disabled, BlockedByPolicy, ExcelNotFound }

public sealed record VbaTrustStatus(VbaTrustState State, string? OfficeVersion);

/// <summary>
/// Reads and, only when the user asks, enables "Trust access to the VBA project object model"
/// for the current Windows user (HKCU). Group Policy values are respected and never overwritten.
/// </summary>
public sealed class ExcelTrustSettings(IRegistryStore? registry = null)
{
    private const string ValueName = "AccessVBOM";
    private static readonly Regex CurVer = new(@"^Excel\.Application\.(\d+)$", RegexOptions.CultureInvariant);
    private readonly IRegistryStore _registry = registry ?? new WindowsRegistryStore();

    public VbaTrustStatus GetState()
    {
        var version = OfficeVersion();
        if (version is null) return new(VbaTrustState.ExcelNotFound, null);

        var policy = _registry.ReadDword(RegistryRoot.CurrentUser, PolicyKey(version), ValueName)
                     ?? _registry.ReadDword(RegistryRoot.LocalMachine, PolicyKey(version), ValueName);
        if (policy is not null)
            return new(policy == 1 ? VbaTrustState.Enabled : VbaTrustState.BlockedByPolicy, version);

        var user = _registry.ReadDword(RegistryRoot.CurrentUser, UserKey(version), ValueName)
                   ?? _registry.ReadDword(RegistryRoot.LocalMachine, UserKey(version), ValueName);
        return new(user == 1 ? VbaTrustState.Enabled : VbaTrustState.Disabled, version);
    }

    public VbaTrustStatus TryEnableForCurrentUser()
    {
        var state = GetState();
        if (state.State != VbaTrustState.Disabled) return state;
        _registry.WriteDword(RegistryRoot.CurrentUser, UserKey(state.OfficeVersion!), ValueName, 1);
        return GetState();
    }

    private string? OfficeVersion()
    {
        var curVer = _registry.ReadString(RegistryRoot.ClassesRoot, @"Excel.Application\CurVer", "");
        var match = curVer is null ? Match.Empty : CurVer.Match(curVer);
        return match.Success ? match.Groups[1].Value + ".0" : null;
    }

    private static string UserKey(string version) => $@"Software\Microsoft\Office\{version}\Excel\Security";
    private static string PolicyKey(string version) => $@"Software\Policies\Microsoft\Office\{version}\Excel\Security";
}

/// <summary>Registry access for Windows; returns nothing on other platforms.</summary>
public sealed class WindowsRegistryStore : IRegistryStore
{
    public string? ReadString(RegistryRoot root, string subKey, string name) =>
        Read(root, subKey, name) as string;

    public int? ReadDword(RegistryRoot root, string subKey, string name) =>
        Read(root, subKey, name) is int value ? value : null;

    public void WriteDword(RegistryRoot root, string subKey, string name, int value)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("Chỉ hỗ trợ trên Windows.");
        using var key = Open(root).CreateSubKey(subKey, writable: true);
        key.SetValue(name, value, RegistryValueKind.DWord);
    }

    private static object? Read(RegistryRoot root, string subKey, string name)
    {
        if (!OperatingSystem.IsWindows()) return null;
        try
        {
            using var key = Open(root).OpenSubKey(subKey);
            return key?.GetValue(name);
        }
        catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException)
        {
            return null;
        }
    }

    [System.Runtime.Versioning.SupportedOSPlatform("windows")]
    private static RegistryKey Open(RegistryRoot root) => root switch
    {
        RegistryRoot.CurrentUser => Registry.CurrentUser,
        RegistryRoot.LocalMachine => Registry.LocalMachine,
        _ => Registry.ClassesRoot
    };
}
