using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class ExcelTrustSettingsTests
{
    private const string UserKey = @"Software\Microsoft\Office\16.0\Excel\Security";
    private const string PolicyKey = @"Software\Policies\Microsoft\Office\16.0\Excel\Security";

    [Fact]
    public void Reports_excel_not_found_without_cur_ver()
    {
        var settings = new ExcelTrustSettings(new FakeRegistry());

        Assert.Equal(VbaTrustState.ExcelNotFound, settings.GetState().State);
    }

    [Fact]
    public void Reports_disabled_then_enabled_after_enabling_for_current_user()
    {
        var registry = FakeRegistry.WithExcel16();
        var settings = new ExcelTrustSettings(registry);

        Assert.Equal(VbaTrustState.Disabled, settings.GetState().State);
        var after = settings.TryEnableForCurrentUser();

        Assert.Equal(VbaTrustState.Enabled, after.State);
        Assert.Equal(1, registry.Dwords[(RegistryRoot.CurrentUser, UserKey, "AccessVBOM")]);
        Assert.Equal("16.0", after.OfficeVersion);
    }

    [Fact]
    public void Does_not_write_when_group_policy_blocks_access()
    {
        var registry = FakeRegistry.WithExcel16();
        registry.Dwords[(RegistryRoot.LocalMachine, PolicyKey, "AccessVBOM")] = 0;
        var settings = new ExcelTrustSettings(registry);

        var result = settings.TryEnableForCurrentUser();

        Assert.Equal(VbaTrustState.BlockedByPolicy, result.State);
        Assert.False(registry.Dwords.ContainsKey((RegistryRoot.CurrentUser, UserKey, "AccessVBOM")));
    }

    [Fact]
    public void Policy_that_allows_access_counts_as_enabled()
    {
        var registry = FakeRegistry.WithExcel16();
        registry.Dwords[(RegistryRoot.CurrentUser, PolicyKey, "AccessVBOM")] = 1;

        Assert.Equal(VbaTrustState.Enabled, new ExcelTrustSettings(registry).GetState().State);
    }

    private sealed class FakeRegistry : IRegistryStore
    {
        public Dictionary<(RegistryRoot, string, string), string> Strings { get; } = [];
        public Dictionary<(RegistryRoot, string, string), int> Dwords { get; } = [];

        public static FakeRegistry WithExcel16()
        {
            var registry = new FakeRegistry();
            registry.Strings[(RegistryRoot.ClassesRoot, @"Excel.Application\CurVer", "")] = "Excel.Application.16";
            return registry;
        }

        public string? ReadString(RegistryRoot root, string subKey, string name) =>
            Strings.TryGetValue((root, subKey, name), out var value) ? value : null;

        public int? ReadDword(RegistryRoot root, string subKey, string name) =>
            Dwords.TryGetValue((root, subKey, name), out var value) ? value : null;

        public void WriteDword(RegistryRoot root, string subKey, string name, int value) =>
            Dwords[(root, subKey, name)] = value;
    }
}
