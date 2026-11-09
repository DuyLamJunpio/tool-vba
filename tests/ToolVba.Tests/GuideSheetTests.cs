using ToolVba.Core;

namespace ToolVba.Tests;

public sealed class GuideSheetTests
{
    private static WorkbookSheetInfo Sheet(string name, SheetVisibility visibility = SheetVisibility.Visible, bool isChart = false) =>
        new(name, isChart, visibility);

    [Fact]
    public void Adds_the_guide_sheet_and_protects_every_other_sheet_in_order()
    {
        var (plan, error) = GuideSheet.Plan([Sheet("Dữ liệu"), Sheet("Báo cáo", SheetVisibility.Hidden), Sheet("Tham số", SheetVisibility.VeryHidden)],
            structureProtected: false, guideMadeByTool: false);

        Assert.Null(error);
        Assert.True(plan!.CreateGuide);
        Assert.Equal(["Dữ liệu", "Báo cáo", "Tham số"], plan.ProtectedNames);
    }

    [Fact]
    public void Reuses_the_guide_sheet_of_a_file_this_tool_already_packaged()
    {
        var (plan, error) = GuideSheet.Plan([Sheet("huong dan"), Sheet("Dữ liệu", SheetVisibility.VeryHidden)],
            structureProtected: false, guideMadeByTool: true);

        Assert.Null(error);
        Assert.False(plan!.CreateGuide);
        Assert.Equal(["Dữ liệu"], plan.ProtectedNames);
    }

    [Fact]
    public void Stops_when_the_product_has_its_own_sheet_named_like_the_guide()
    {
        var (plan, error) = GuideSheet.Plan([Sheet("Huong dan"), Sheet("Dữ liệu")], structureProtected: false, guideMadeByTool: false);

        Assert.Null(plan);
        Assert.Contains("Huong dan", error);
    }

    [Fact]
    public void Stops_on_a_chart_sheet_because_the_module_hides_worksheets_only()
    {
        var (plan, error) = GuideSheet.Plan([Sheet("Dữ liệu"), Sheet("Biểu đồ", isChart: true)], structureProtected: false, guideMadeByTool: false);

        Assert.Null(plan);
        Assert.Contains("Biểu đồ", error);
    }

    [Fact]
    public void Stops_when_the_workbook_structure_is_protected()
    {
        var (plan, error) = GuideSheet.Plan([Sheet("Dữ liệu")], structureProtected: true, guideMadeByTool: false);

        Assert.Null(plan);
        Assert.Contains("cấu trúc", error);
    }

    [Fact]
    public void Stops_when_nothing_but_the_guide_would_be_left_to_protect()
    {
        var (plan, error) = GuideSheet.Plan([Sheet("Huong dan")], structureProtected: false, guideMadeByTool: true);

        Assert.Null(plan);
        Assert.NotNull(error);
    }

    [Fact]
    public void Guide_text_explains_macros_unblock_the_copied_code_and_the_store_page()
    {
        var text = string.Join("\n", GuideSheet.Lines);

        Assert.Contains("Enable Content", text);
        Assert.Contains("Unblock", text);
        Assert.Contains("Properties", text);
        Assert.Contains("Ctrl+V", text);
        Assert.Contains("https://shop.sunprime.vn/guides/macros", text);
    }
}
