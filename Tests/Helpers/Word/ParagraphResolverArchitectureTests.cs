using System.Text;

namespace AsposeMcpServer.Tests.Helpers.Word;

/// <summary>
///     Architecture guard (L1) for the unified paragraph-addressing scheme. Every Word handler must
///     locate paragraphs by index through <c>ParagraphResolver</c>, so the index space stays
///     consistent across all tools. Raw <c>GetChildNodes(NodeType.Paragraph, ...)</c> access is only
///     legitimate for non-resolution purposes — counting, iterating every paragraph, scanning/search,
///     append-at-end, building a base list for a get operation, or working inside an already-resolved
///     container (a cell, text box, or section body).
/// </summary>
/// <remarks>
///     Each production handler under Handlers/Word that still references the raw pattern has been
///     individually audited and listed in <see cref="Allowlist" /> with its justification. A handler
///     NOT on the allowlist that uses the pattern fails this test: either route its index resolution
///     through ParagraphResolver, or — if the use is one of the legitimate non-resolution cases —
///     audit it and add it to the allowlist.
/// </remarks>
public class ParagraphResolverArchitectureTests
{
    private const string RawParagraphAccessPattern = "GetChildNodes(NodeType.Paragraph";

    /// <summary>
    ///     Production Handlers/Word files permitted to use raw paragraph access, each with the reason
    ///     the use is not index resolution. Paths are relative to the Handlers/Word root, '/'-separated.
    /// </summary>
    private static readonly Dictionary<string, string> Allowlist = new()
    {
        ["Content/GetWordDocumentInfoHandler.cs"] =
            "iterates each section body for tab stops (already section-relative)",
        ["List/AddWordListHandler.cs"] = "scans backward for the most recent list item",
        ["List/DeleteWordListItemHandler.cs"] = "reports remaining paragraph count",
        ["List/GetWordListFormatHandler.cs"] = "builds the global list/dict; resolution goes through the resolver",
        ["List/RestartWordListNumberingHandler.cs"] = "iterates the contiguous run from a resolver-resolved start",
        ["Page/AddBlankPageWordHandler.cs"] = "searches by page index via LayoutCollector, not paragraph index",
        ["Paragraph/DeleteParagraphWordHandler.cs"] = "reports remaining paragraph count",
        ["Paragraph/GetParagraphsWordHandler.cs"] = "builds the base list to enumerate; emits resolver addresses",
        ["Paragraph/InsertParagraphWordHandler.cs"] = "reports total paragraph count",
        ["Paragraph/MergeParagraphsWordHandler.cs"] = "reports remaining paragraph count",
        ["Shape/EditTextBoxContentWordHandler.cs"] = "operates within an already-resolved text box",
        ["Styles/ApplyWordStyleHandler.cs"] = "applies a style to every paragraph (apply-to-all)",
        ["Styles/ListWordStylesHandler.cs"] = "scans every paragraph to collect used style names",
        ["Table/CopyWordTableHandler.cs"] = "reads the target section body for table placement",
        ["Table/EditCellFormatWordTableHandler.cs"] = "operates within an already-resolved table cell",
        ["Text/AddWithStyleWordTextHandler.cs"] = "anchors within a resolver-resolved parent node",
        ["Text/AddWordTextHandler.cs"] = "moves to the last body paragraph for append-at-end",
        ["Text/DeleteWordTextHandler.cs"] = "global deletion canvas + searchText scan; index path uses the resolver",
        ["Text/SearchWordTextHandler.cs"] = "iterates every paragraph to search; emits resolver addresses"
    };

    private static readonly string WordHandlersRoot = ResolveWordHandlersRoot();

    /// <summary>
    ///     Fails if any non-allowlisted Handlers/Word source file uses raw paragraph access, which is
    ///     almost always an index-resolution path that should go through ParagraphResolver instead.
    /// </summary>
    [Fact]
    public void WordHandlers_ResolveParagraphsThroughResolver_NotRawGetChildNodes()
    {
        Assert.True(Directory.Exists(WordHandlersRoot),
            $"Handlers/Word root not found: '{WordHandlersRoot}'. The test cannot locate the handler sources.");

        var files = Directory.GetFiles(WordHandlersRoot, "*.cs", SearchOption.AllDirectories)
            .OrderBy(f => f)
            .ToArray();
        Assert.NotEmpty(files);

        var offenders = new List<string>();
        foreach (var filePath in files)
        {
            var source = File.ReadAllText(filePath, Encoding.UTF8);
            if (!source.Contains(RawParagraphAccessPattern, StringComparison.Ordinal))
                continue;

            var relative = Path.GetRelativePath(WordHandlersRoot, filePath).Replace('\\', '/');
            if (!Allowlist.ContainsKey(relative))
                offenders.Add(relative);
        }

        if (offenders.Count > 0)
            Assert.Fail(BuildOffenderMessage(offenders));
    }

    /// <summary>
    ///     Keeps the allowlist honest: every allowlisted path must still exist and still use the raw
    ///     pattern, so renamed, deleted, or fully-migrated entries are pruned rather than left stale.
    /// </summary>
    [Fact]
    public void Allowlist_HasNoStaleEntries()
    {
        Assert.True(Directory.Exists(WordHandlersRoot),
            $"Handlers/Word root not found: '{WordHandlersRoot}'.");

        var stale = new List<string>();
        foreach (var (relative, _) in Allowlist)
        {
            var fullPath = Path.Combine(WordHandlersRoot, relative.Replace('/', Path.DirectorySeparatorChar));
            if (!File.Exists(fullPath) ||
                !File.ReadAllText(fullPath, Encoding.UTF8)
                    .Contains(RawParagraphAccessPattern, StringComparison.Ordinal))
                stale.Add(relative);
        }

        Assert.True(stale.Count == 0,
            "Stale ParagraphResolver allowlist entries (file missing or no longer uses raw paragraph access — " +
            $"remove them from the allowlist):{Environment.NewLine}  {string.Join($"{Environment.NewLine}  ", stale)}");
    }

    private static string BuildOffenderMessage(List<string> offenders)
    {
        var sb = new StringBuilder();
        sb.AppendLine(
            $"PARAGRAPH ADDRESSING GUARD: {offenders.Count} Handlers/Word file(s) use raw " +
            $"'{RawParagraphAccessPattern}...' access but are not on the audited allowlist:");
        foreach (var offender in offenders)
            sb.AppendLine($"  - {offender}");
        sb.AppendLine();
        sb.AppendLine("Resolve paragraphs by index through ParagraphResolver.Resolve / AddressOf so the index");
        sb.AppendLine("space stays consistent across every tool. If this use is genuinely not resolution");
        sb.AppendLine("(count, iterate-all, scan/search, append-at-end, base list, or within an already-resolved");
        sb.AppendLine("container), audit it and add it to the allowlist with a justification.");
        return sb.ToString();
    }

    /// <summary>
    ///     Resolves the absolute path to the production Handlers/Word directory by walking up from the
    ///     test assembly output until a directory is found that contains both "Handlers" and the main
    ///     project file "AsposeMcpServer.csproj".
    /// </summary>
    private static string ResolveWordHandlersRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var handlers = Path.Combine(dir.FullName, "Handlers");
            var csproj = Path.Combine(dir.FullName, "AsposeMcpServer.csproj");
            if (Directory.Exists(handlers) && File.Exists(csproj))
                return Path.Combine(handlers, "Word");
            dir = dir.Parent;
        }

        return Path.Combine(AppContext.BaseDirectory, "Handlers", "Word");
    }
}
