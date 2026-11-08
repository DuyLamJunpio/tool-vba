using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class OutputPathResolverTests
{
    [Fact]
    public void Places_macro_enabled_output_next_to_source()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("Báo cáo tháng.xlsx");
        File.WriteAllText(input, "");

        Assert.Equal(temp.PathFor("Báo cáo tháng_vba.xlsm"), OutputPathResolver.Resolve(input));
    }

    [Fact]
    public void Numbers_output_when_name_is_taken_and_never_returns_the_source()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("Report_vba.xlsm");
        File.WriteAllText(input, "");
        File.WriteAllText(temp.PathFor("Report_vba (2).xlsm"), "");

        var output = OutputPathResolver.Resolve(input);

        Assert.Equal(temp.PathFor("Report_vba (3).xlsm"), output);
    }

    [Fact]
    public void Skips_names_reserved_by_other_files_in_the_same_batch()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("a.xlsm");
        var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { temp.PathFor("a_vba.xlsm") };

        Assert.Equal(temp.PathFor("a_vba (2).xlsm"), OutputPathResolver.Resolve(input, reserved: reserved));
    }

    [Fact]
    public void Keeps_numbers_in_parentheses_that_belong_to_the_original_name()
    {
        using var temp = new TestDirectory();

        Assert.Equal(temp.PathFor("Report (2024)_vba.xlsm"), OutputPathResolver.Resolve(temp.PathFor("Report (2024).xlsx")));
    }

    [Fact]
    public void Uses_chosen_directory_when_given()
    {
        using var temp = new TestDirectory();
        var input = temp.PathFor("a.xlsx");
        var other = Directory.CreateDirectory(temp.PathFor("out")).FullName;

        Assert.Equal(Path.Combine(other, "a_vba.xlsm"), OutputPathResolver.Resolve(input, other));
    }
}
