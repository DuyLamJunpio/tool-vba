namespace ToolVba.Core;

/// <summary>
/// Remembers which components a workbook received from the company package, using custom document
/// properties. Excel limits a text property to 255 characters, so the list is split into numbered chunks.
/// </summary>
public static class ManagedComponentList
{
    public const int MaxChunkLength = 240;
    public const string ChunkPropertyPrefix = "ToolVba.Managed.";
    public const string VersionProperty = "ToolVba.PackageVersion";
    private const int MaxChunks = 64;
    private const char Separator = ';';

    public static IReadOnlyList<string> ToChunks(IEnumerable<string> names)
    {
        var chunks = new List<string>();
        var current = "";
        foreach (var name in names)
        {
            var candidate = current.Length == 0 ? name : current + Separator + name;
            if (candidate.Length <= MaxChunkLength)
            {
                current = candidate;
                continue;
            }
            chunks.Add(current);
            current = name;
        }
        if (current.Length > 0) chunks.Add(current);
        return chunks;
    }

    public static IReadOnlyList<string> FromChunks(IEnumerable<string> chunks) =>
        chunks.SelectMany(c => c.Split(Separator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)).ToList();

    internal static IReadOnlyList<string> Read(IVbaWorkbookSession session) => FromChunks(ReadChunks(session));

    internal static void Write(IVbaWorkbookSession session, IEnumerable<string> names, string packageVersion)
    {
        var previousCount = ReadChunks(session).Count;
        var chunks = ToChunks(names);
        for (var i = 0; i < chunks.Count; i++)
            session.SetCustomProperty(ChunkPropertyPrefix + (i + 1), chunks[i]);
        for (var i = chunks.Count + 1; i <= previousCount; i++)
            session.DeleteCustomProperty(ChunkPropertyPrefix + i);
        session.SetCustomProperty(VersionProperty, packageVersion);
    }

    private static List<string> ReadChunks(IVbaWorkbookSession session)
    {
        var chunks = new List<string>();
        for (var i = 1; i <= MaxChunks; i++)
        {
            var value = session.GetCustomProperty(ChunkPropertyPrefix + i);
            if (string.IsNullOrEmpty(value)) break;
            chunks.Add(value);
        }
        return chunks;
    }
}
