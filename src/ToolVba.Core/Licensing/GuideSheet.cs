namespace ToolVba.Core;

public enum SheetVisibility { Visible, Hidden, VeryHidden }

/// <param name="IsChart">A chart sheet: the licensing module hides worksheets only.</param>
public sealed record WorkbookSheetInfo(string Name, bool IsChart, SheetVisibility Visibility);

/// <param name="CreateGuide">False when the guide sheet of an earlier packaging is reused.</param>
/// <param name="ProtectedNames">Every sheet except the guide, in workbook order: saved very hidden.</param>
public sealed record SheetPlan(bool CreateGuide, IReadOnlyList<string> ProtectedNames);

/// <summary>
/// The guide sheet a product shows before its computer is activated (DEC-112 items 2 and 7): it is the
/// only visible sheet of the packaged file, and every other sheet is protected by the licensing module.
/// </summary>
public static class GuideSheet
{
    public const string Name = "Huong dan";

    /// <summary>Document property naming the guide sheet this tool created, so packaging again reuses it.</summary>
    public const string MarkerProperty = "ToolVba.GuideSheet";

    public static IReadOnlyList<string> Lines { get; } =
    [
        "File này cần bật macro để sử dụng.",
        "1. Bấm Enable Content (Bật nội dung) trên thanh màu vàng.",
        "2. Bấm Yes ở hộp thoại kích hoạt. Mã kích hoạt đã được sao chép sẵn: trên trang web chỉ cần bấm Ctrl+V rồi xác nhận.",
        "",
        "Nếu Excel hiện thanh màu đỏ \"SECURITY RISK\" (Rủi ro bảo mật) và không có nút Enable Content:",
        "  a. Đóng file này lại.",
        "  b. Chuột phải vào file > Properties (Thuộc tính) > thẻ General: đánh dấu ô Unblock (Bỏ chặn) > OK.",
        "  c. Mở lại file và bấm Enable Content. Mỗi file chỉ cần làm một lần.",
        "",
        "Hướng dẫn có hình: https://shop.sunprime.vn/guides/macros"
    ];

    public static (SheetPlan? Plan, string? Error) Plan(
        IReadOnlyList<WorkbookSheetInfo> sheets, bool structureProtected, bool guideMadeByTool)
    {
        if (structureProtected)
            return (null, "Workbook đang khóa cấu trúc (Review > Protect Workbook) nên không thêm hay ẩn sheet được. " +
                          "Hãy bỏ khóa cấu trúc trong file gốc rồi thử lại.");
        var chart = sheets.FirstOrDefault(s => s.IsChart);
        if (chart is not null)
            return (null, $"File có chart sheet \"{chart.Name}\". Module bản quyền chỉ ẩn được worksheet; " +
                          "hãy chuyển biểu đồ vào một worksheet rồi thử lại.");
        var guide = sheets.FirstOrDefault(s => string.Equals(s.Name, Name, StringComparison.OrdinalIgnoreCase));
        if (guide is not null && !guideMadeByTool)
            return (null, $"File đã có sheet \"{guide.Name}\" không do tool-vba tạo. Hãy đổi tên sheet đó rồi thử lại.");
        var protectedNames = sheets.Where(s => !ReferenceEquals(s, guide)).Select(s => s.Name).ToList();
        if (protectedNames.Count == 0)
            return (null, "File không có sheet nào ngoài sheet hướng dẫn để bảo vệ.");
        return (new SheetPlan(guide is null, protectedNames), null);
    }
}
