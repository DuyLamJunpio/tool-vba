using System.Text.RegularExpressions;

namespace ToolVba.Core;

public sealed record DocumentMergeResult(bool Succeeded, string? Code, IReadOnlyList<string> ConflictingNames, string? Error)
{
    public static DocumentMergeResult Conflict(IReadOnlyList<string> names) =>
        new(false, null, names, "Trùng tên với code có sẵn: " + string.Join(", ", names) + ".");

    public static DocumentMergeResult Failure(string error) => new(false, null, [], error);
}

/// <summary>
/// Merges company code into a document module (ThisWorkbook or a sheet) that may already hold employee code.
/// Company declarations go into a marked region after the existing declarations; company procedures go
/// into a marked region at the end. Earlier managed regions are replaced, never duplicated.
/// The licensing events (DEC-115) are the exception: when the product already handles one, the package's
/// body goes at the top of that handler between hook markers.
/// Name detection is a best-effort line scan: it covers the common forms that would otherwise cause an
/// "Ambiguous name detected" compile error.
/// </summary>
public static class DocumentModuleMerger
{
    public const string DeclarationsBegin = "'=== tool-vba BEGIN declarations (auto-managed, do not edit) ===";
    public const string DeclarationsEnd = "'=== tool-vba END declarations ===";
    public const string CodeBegin = "'=== tool-vba BEGIN code (auto-managed, do not edit) ===";
    public const string CodeEnd = "'=== tool-vba END code ===";
    public const string HookBegin = "'=== tool-vba BEGIN license hook (auto-managed, do not edit) ===";
    public const string HookEnd = "'=== tool-vba END license hook ===";

    private const RegexOptions Options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;
    private static readonly Regex ProcedureStart = new(
        @"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)", Options);
    private static readonly Regex OptionLine = new(@"^\s*Option\s+(Explicit|Base|Compare|Private\s+Module)\b", Options);
    private static readonly Regex[] DeclarationNames =
    [
        new(@"^\s*(?:(?:Public|Private|Global)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+([A-Za-z_]\w*)", Options),
        new(@"^\s*(?:(?:Public|Private)\s+)?(?:Type|Enum|Event)\s+([A-Za-z_]\w*)", Options),
        new(@"^\s*(?:(?:Public|Private|Global)\s+)?Const\s+([A-Za-z_]\w*)", Options),
        new(@"^\s*(?:Dim|Public|Private|Global)\s+(?:WithEvents\s+)?(?!(?:Declare|Type|Enum|Event|Const|Sub|Function|Property|Static)\b)([A-Za-z_]\w*)", Options)
    ];

    /// <summary>The licensing events whose body is put into the product's own handler (DEC-115).</summary>
    private static readonly HashSet<string> HookableEvents = new(StringComparer.OrdinalIgnoreCase)
    {
        "Workbook_BeforeClose", "Workbook_BeforeSave", "Workbook_AfterSave"
    };

    /// <summary>A whole Sub signature on one line, with nothing after the parameter list.</summary>
    private static readonly Regex SubSignature = new(
        @"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?Sub\s+([A-Za-z_][A-Za-z0-9_]*)\s*(\([^()]*\))\s*$", Options);
    private static readonly Regex EndSub = new(@"^\s*End\s+Sub\b", Options);

    /// <summary>
    /// A hooked body runs inside the product's handler: a local declaration could clash with the product's
    /// locals, and an error handler, label or jump would change how the product's own lines run.
    /// </summary>
    private static readonly Regex UnsafeHookLine = new(
        @"^\s*(?:(?:Dim|Static|ReDim|Const)\b|On\s+Error\b|[A-Za-z_]\w*:\s*$)|\b(?:GoTo|Resume)\b", Options);

    /// <summary>
    /// The close call goes last (owner, 08/11/2026): it runs only when the product did not cancel, so a product
    /// that asks "close?" and is answered No keeps its sheets and access.
    /// </summary>
    private const string CloseEvent = "Workbook_BeforeClose";

    /// <summary>"If cond Then stmt" on one line without Else (a line with Else does not match).</summary>
    private static readonly Regex SingleLineIf = new(@"^If\s+(?!.*\bElse\b)(.+?)\s+Then\s+(\S.*)$", Options);

    /// <summary>A statement that leaves the handler before its last line: Exit Sub/Function or a bare End.</summary>
    private static readonly Regex EarlyExit = new(
        @"\bExit\s+(?:Sub|Function)\b|(?:^|:|\bThen\b|\bElse\b)\s*End\s*(?:'.*)?$", Options);

    public static DocumentMergeResult Merge(string existingCode, string packageCode)
    {
        var withoutRegions = RemoveManagedRegions(SplitLines(existingCode));
        if (withoutRegions is null)
            return DocumentMergeResult.Failure("Vùng code tool-vba trong module thiếu dòng END nên không thể thay an toàn. " +
                                               "Hãy mở Alt+F11, xóa vùng code tool-vba bị sửa dở rồi thử lại.");
        var userLines = TrimBlankEdges(withoutRegions);
        // A template that was itself produced by this tool still carries marker lines; they must not nest.
        var package = SplitLines(StripMarkers(packageCode))
            .Where(l => !l.TrimStart().StartsWith("Attribute VB_", StringComparison.OrdinalIgnoreCase)).ToList();

        var packageSplit = FirstProcedure(package);
        var packageOptions = package.Take(packageSplit).Where(IsOption).ToList();
        var packageDeclarations = TrimBlankEdges(package.Take(packageSplit).Where(l => !IsOption(l)).ToList());
        var (user, packageProcedures, hookError) = ApplyLicenseHooks(userLines, TrimBlankEdges(package.Skip(packageSplit).ToList()));
        if (hookError is not null) return DocumentMergeResult.Failure(hookError);

        var userSplit = FirstProcedure(user);
        var conflicts = Names(packageProcedures, ProcedureNames).Concat(Names(packageDeclarations, DeclarationNamesOf))
            .Intersect(Names(user, ProcedureNames).Concat(Names(user.Take(userSplit).ToList(), DeclarationNamesOf)), StringComparer.OrdinalIgnoreCase)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (conflicts.Count > 0) return DocumentMergeResult.Conflict(conflicts);

        var (missingOptions, optionError) = MissingOptions(packageOptions, user.Take(userSplit).Where(IsOption).ToList(), user);
        if (optionError is not null) return DocumentMergeResult.Failure(optionError);

        var result = new List<string>(missingOptions);
        result.AddRange(TrimBlankEdges(user.Take(userSplit).ToList()));
        if (packageDeclarations.Count > 0)
            AppendRegion(result, DeclarationsBegin, packageDeclarations, DeclarationsEnd);
        var userProcedures = TrimBlankEdges(user.Skip(userSplit).ToList());
        if (userProcedures.Count > 0)
        {
            if (result.Count > 0) result.Add("");
            result.AddRange(userProcedures);
        }
        packageProcedures = TrimBlankEdges(packageProcedures);
        if (packageProcedures.Count > 0)
            AppendRegion(result, CodeBegin, packageProcedures, CodeEnd);
        return new(true, string.Join("\r\n", result) + "\r\n", [], null);
    }

    public static bool ContainsManagedCode(string code) =>
        SplitLines(code).Any(l => IsMarker(l, DeclarationsBegin) || IsMarker(l, CodeBegin) || IsMarker(l, HookBegin));

    /// <summary>Removes only the marker lines, keeping the code between them.</summary>
    public static string StripMarkers(string code) =>
        string.Join("\r\n", SplitLines(code).Where(l => !IsAnyMarker(l)));

    /// <summary>True when a module has nothing but blank lines, comments and Option statements.</summary>
    public static bool IsEffectivelyEmpty(string code) =>
        SplitLines(code).All(l => string.IsNullOrWhiteSpace(l) || IsOption(l) || l.TrimStart().StartsWith('\''));

    /// <summary>
    /// Moves the body of each licensing event the product already handles to the top of the product's
    /// handler, and drops it from the package procedures. The two signatures must match (case, spacing
    /// and scope aside); anything else stops with a message naming the handler.
    /// </summary>
    private static (List<string> User, List<string> PackageProcedures, string? Error) ApplyLicenseHooks(
        List<string> user, List<string> packageProcedures)
    {
        var result = new List<string>(user);
        var remaining = new List<string>(packageProcedures);
        foreach (var name in Names(packageProcedures, ProcedureNames).Where(HookableEvents.Contains).ToList())
        {
            var userStart = result.FindIndex(l => string.Equals(ProcedureNames(l), name, StringComparison.OrdinalIgnoreCase));
            if (userStart < 0) continue;
            var packageStart = remaining.FindIndex(l => string.Equals(ProcedureNames(l), name, StringComparison.OrdinalIgnoreCase));
            var packageEnd = remaining.FindIndex(packageStart + 1, l => EndSub.IsMatch(l));
            var userSignature = NormalizedSignature(result[userStart]);
            if (packageEnd < 0 || userSignature is null || userSignature != NormalizedSignature(remaining[packageStart]))
                return ([], [], $"Thủ tục {name} trong file có chữ ký khác chữ ký chuẩn của Excel ({result[userStart].Trim()}) " +
                                "nên không ghép được code bản quyền. Hãy sửa dòng khai báo cho giống chữ ký chuẩn (viết trên một dòng) rồi thử lại.");
            var body = remaining.GetRange(packageStart + 1, packageEnd - packageStart - 1);
            if (body.Any(l => !IsComment(l) && UnsafeHookLine.IsMatch(l)))
                return ([], [], $"Code sự kiện {name} của gói có khai báo, nhãn hoặc xử lý lỗi nên không ghép vào thủ tục có sẵn của file được. " +
                                "Thân sự kiện bản quyền chỉ được gồm lời gọi module bản quyền; hãy sửa file mẫu rồi xuất bản lại gói.");
            var insertError = InsertHook(result, userStart, name, body);
            if (insertError is not null) return ([], [], insertError);
            remaining.RemoveRange(packageStart, packageEnd - packageStart + 1);
        }
        return (result, remaining, null);
    }

    /// <summary>At the top of the handler, except BeforeClose: last, and only when the product did not cancel.</summary>
    private static string? InsertHook(List<string> user, int userStart, string name, List<string> body)
    {
        if (!string.Equals(name, CloseEvent, StringComparison.OrdinalIgnoreCase))
        {
            user.InsertRange(userStart + 1, [HookBegin, .. body, HookEnd]);
            return null;
        }
        var userEnd = user.FindIndex(userStart + 1, l => EndSub.IsMatch(l));
        if (userEnd < 0)
            return $"Thủ tục {name} trong file thiếu dòng End Sub nên không ghép được code bản quyền.";
        if (user.GetRange(userStart + 1, userEnd - userStart - 1).Any(l => !IsComment(l) && EarlyExit.IsMatch(l)))
            return $"Thủ tục {name} trong file có Exit Sub, Exit Function hoặc End, nên lời gọi bản quyền đặt ở cuối thủ tục có thể không chạy. " +
                   "Hãy viết lại thủ tục không dùng các lệnh này (ví dụ dùng If … Else … End If) rồi thử lại.";
        user.InsertRange(userEnd, [HookBegin, .. GuardedByCancel(body), HookEnd]);
        return null;
    }

    /// <summary>
    /// "If Not Cancel Then …" on one line for one statement, an If block otherwise. A single-line
    /// "If cond Then stmt" (the shipped "If ThisWorkbook.Saved Then …", owner 08/11/2026 after CL-1) becomes
    /// "If Not Cancel And (cond) Then stmt".
    /// </summary>
    private static List<string> GuardedByCancel(List<string> body)
    {
        // A comment is no statement: "If Not Cancel Then ' note" would open a block If without End If.
        var statements = body.Where(l => !string.IsNullOrWhiteSpace(l) && !IsComment(l)).Select(l => l.Trim()).ToList();
        if (statements.Count != 1) return ["    If Not Cancel Then", .. statements.Select(s => "        " + s), "    End If"];
        var singleLineIf = SingleLineIf.Match(statements[0]);
        if (singleLineIf.Success)
            return [$"    If Not Cancel And ({singleLineIf.Groups[1].Value}) Then {singleLineIf.Groups[2].Value}"];
        // An If with Else on one line cannot take another condition without changing what Else covers.
        if (statements[0].StartsWith("If ", StringComparison.OrdinalIgnoreCase))
            return ["    If Not Cancel Then", "        " + statements[0], "    End If"];
        return ["    If Not Cancel Then " + statements[0]];
    }

    private static bool IsComment(string line) => line.TrimStart().StartsWith('\'');

    /// <summary>"name(parameters)" without scope, case or extra spaces; null when not a one-line Sub signature.</summary>
    private static string? NormalizedSignature(string line)
    {
        var match = SubSignature.Match(line);
        if (!match.Success) return null;
        var parameters = Regex.Replace(match.Groups[2].Value, @"\s+", " ");
        parameters = Regex.Replace(parameters, @"\s*([(),])\s*", "$1");
        return (match.Groups[1].Value + parameters).ToLowerInvariant();
    }

    /// <summary>
    /// Option statements of the package that the module still needs. Option Explicit is never required
    /// (code that compiles with it also compiles without it); a different value of the same option
    /// (Compare Text vs Binary, Base 0 vs 1) would change existing code, so it is refused.
    /// </summary>
    private static (List<string> Missing, string? Error) MissingOptions(
        IReadOnlyList<string> packageOptions, IReadOnlyList<string> userOptions, IReadOnlyList<string> user)
    {
        var missing = new List<string>();
        foreach (var option in packageOptions)
        {
            var keyword = OptionKeyword(option);
            if (keyword == "explicit") continue;
            var existing = userOptions.FirstOrDefault(u => OptionKeyword(u) == keyword);
            if (existing is null) missing.Add(option.Trim());
            else if (!Normalize(existing).Equals(Normalize(option), StringComparison.OrdinalIgnoreCase))
                return ([], $"Module đang dùng '{existing.Trim()}' nhưng code công ty cần '{option.Trim()}'. Hãy thống nhất dòng Option này.");
        }
        var userHasCode = user.Any(l => !string.IsNullOrWhiteSpace(l) && !IsOption(l) && !l.TrimStart().StartsWith('\''));
        return missing.Count > 0 && userHasCode
            ? ([], $"Code công ty cần '{missing[0]}' nhưng module đã có code khác. Hãy thêm dòng này vào file Excel hoặc bỏ khỏi code mẫu.")
            : (missing, null);
    }

    private static string OptionKeyword(string line)
    {
        var match = OptionLine.Match(line);
        return match.Success ? Regex.Replace(match.Groups[1].Value, @"\s+", " ").ToLowerInvariant() : "";
    }

    /// <summary>Removes earlier managed regions and hooks; null when one has no END marker, so nothing is lost silently.</summary>
    private static List<string>? RemoveManagedRegions(IEnumerable<string> lines)
    {
        var kept = new List<string>();
        string? closing = null;
        foreach (var line in lines)
        {
            if (closing is not null)
            {
                if (IsMarker(line, closing)) closing = null;
                continue;
            }
            if (IsMarker(line, DeclarationsBegin)) closing = DeclarationsEnd;
            else if (IsMarker(line, CodeBegin)) closing = CodeEnd;
            else if (IsMarker(line, HookBegin)) closing = HookEnd;
            else kept.Add(line);
        }
        return closing is null ? kept : null;
    }

    private static void AppendRegion(List<string> target, string begin, IEnumerable<string> body, string end)
    {
        if (target.Count > 0) target.Add("");
        target.Add(begin);
        target.AddRange(body);
        target.Add(end);
    }

    private static bool IsAnyMarker(string line) =>
        IsMarker(line, DeclarationsBegin) || IsMarker(line, DeclarationsEnd) || IsMarker(line, CodeBegin) ||
        IsMarker(line, CodeEnd) || IsMarker(line, HookBegin) || IsMarker(line, HookEnd);

    private static bool IsMarker(string line, string marker) => line.Trim().Equals(marker, StringComparison.OrdinalIgnoreCase);
    private static bool IsOption(string line) => OptionLine.IsMatch(line);
    private static string Normalize(string line) => Regex.Replace(line.Trim(), @"\s+", " ");
    private static List<string> SplitLines(string code) => [.. VbaSourceRules.SplitLines(code)];

    private static int FirstProcedure(IReadOnlyList<string> lines)
    {
        for (var i = 0; i < lines.Count; i++)
            if (ProcedureStart.IsMatch(lines[i])) return i;
        return lines.Count;
    }

    private static IEnumerable<string> Names(IReadOnlyList<string> lines, Func<string, string?> extract) =>
        lines.Select(extract).OfType<string>();

    private static string? ProcedureNames(string line)
    {
        var match = ProcedureStart.Match(line);
        return match.Success ? match.Groups[1].Value : null;
    }

    private static string? DeclarationNamesOf(string line)
    {
        if (line.TrimStart().StartsWith('\'')) return null;
        foreach (var regex in DeclarationNames)
        {
            var match = regex.Match(line);
            if (match.Success) return match.Groups[1].Value;
        }
        return null;
    }

    private static List<string> TrimBlankEdges(List<string> lines)
    {
        var start = lines.FindIndex(l => !string.IsNullOrWhiteSpace(l));
        if (start < 0) return [];
        var end = lines.FindLastIndex(l => !string.IsNullOrWhiteSpace(l));
        return lines.GetRange(start, end - start + 1);
    }
}
