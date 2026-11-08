using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class OutputPathResolverTests
{
    private static readonly LicenseProduct Product = new("test-01", "1.0.0");

    [Fact]
    public void Names_the_output_after_the_product_and_version_next_to_the_source()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("Báo cáo tháng.xlsx");
        File.WriteAllText(input, "");

        Assert.Equal(temp.PathFor("test-01_1.0.0_ban-quyen.xlsm"), OutputPathResolver.Resolve(input, Product));
    }

    [Fact]
    public void Numbers_output_when_name_is_taken()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("source.xlsm");
        File.WriteAllText(temp.PathFor("test-01_1.0.0_ban-quyen.xlsm"), "");
        File.WriteAllText(temp.PathFor("test-01_1.0.0_ban-quyen (2).xlsm"), "");

        Assert.Equal(temp.PathFor("test-01_1.0.0_ban-quyen (3).xlsm"), OutputPathResolver.Resolve(input, Product));
    }

    [Fact]
    public void Never_returns_the_source_when_it_already_has_the_packaged_name()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("test-01_1.0.0_ban-quyen.xlsm");
        File.WriteAllText(input, "");

        Assert.Equal(temp.PathFor("test-01_1.0.0_ban-quyen (2).xlsm"), OutputPathResolver.Resolve(input, Product));
    }

    [Fact]
    public void Uses_chosen_directory_when_given()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("a.xlsx");
        var other = Directory.CreateDirectory(temp.PathFor("out")).FullName;

        Assert.Equal(Path.Combine(other, "test-01_1.0.0_ban-quyen.xlsm"), OutputPathResolver.Resolve(input, Product, other));
    }
}
