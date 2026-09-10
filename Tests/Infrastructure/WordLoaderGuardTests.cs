using System.Text;
using System.Text.RegularExpressions;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Covers A-03's structural half. Aspose.Words picks the format from the file's content, so an
///     HTML or MHT payload named <c>.docx</c> reaches the same loader and makes the library fetch
///     whatever it references. Only <c>DocumentContext</c> installed a resource guard; eight other
///     places called <c>new Document(path)</c> directly, so whether the SSRF policy applied
///     depended on which entry point the caller used. This fails when a new unguarded load appears.
/// </summary>
public class WordLoaderGuardTests
{
    /// <summary>
    ///     A load from a path. <c>new Document()</c> with no argument creates an empty document and
    ///     reads nothing, so it is not a resource-loading site.
    /// </summary>
    // The type can be written unqualified, fully qualified, or through an alias, and a guard that
    // only knew the first spelling would pass a file that loaded documents through either of the
    // others (§18.5.2).
    private static readonly Regex LoadFromPath = new(
        @"new\s+(?:[A-Za-z_][A-Za-z0-9_]*\s*\.\s*)*(?:Document|WordDocument|AsposeDocument)\s*\(\s*(?!\s*\))",
        RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>Where a load is expected to go instead.</summary>
    private static readonly Regex GuardedLoad = new(
        @"GuardedWordLoader\s*\.\s*Load\s*\(", RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>
    ///     Files allowed to construct a document from a path, with the reason. The loader itself
    ///     must, and the copy in the Aspose namespace is a different type entirely.
    /// </summary>
    /// <summary>
    ///     The one file that may construct a Word document from an input, by its
    ///     repository-relative path. A bare file name let any directory's `GuardedWordLoader.cs`
    ///     inherit the exemption (R22-TST02).
    /// </summary>
    private static readonly Dictionary<string, string> Exempt = new(StringComparer.Ordinal)
    {
        ["Helpers/Word/GuardedWordLoader.cs"] = "This is the guarded loader every other site delegates to."
    };

    /// <summary>A file's path relative to the repository root, with forward slashes.</summary>
    /// <param name="path">The absolute path.</param>
    /// <returns>The repository-relative path.</returns>
    private static string RepositoryRelative(string path)
    {
        return Path.GetRelativePath(RepositoryRoot().FullName, path).Replace('\\', '/');
    }

    /// <summary>Walks up to the repository root, identified by the project file.</summary>
    /// <returns>The repository root directory.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir;
    }

    /// <summary>
    ///     Reports every Word document loaded from a path in one file without the resource guard.
    /// </summary>
    /// <param name="source">The file's text.</param>
    /// <param name="label">How to name the file in the report.</param>
    /// <returns>One entry per unguarded load.</returns>
    internal static List<string> UnguardedLoads(string source, string label)
    {
        var offenders = new List<string>();

        // Aspose.Pdf.Document and Aspose.Words.Document share a type name, and a file that imports
        // Aspose.Pdf without Aspose.Words means the PDF one. Only the Words loader takes a
        // resource-loading callback.
        var importsWords = source.Contains("using Aspose.Words;", StringComparison.Ordinal);
        var importsPdf = source.Contains("using Aspose.Pdf;", StringComparison.Ordinal);

        foreach (Match match in LoadFromPath.Matches(source))
        {
            var lineStart = source.LastIndexOf('\n', Math.Max(match.Index - 1, 0)) + 1;
            var lineEnd = source.IndexOf('\n', match.Index);
            if (lineEnd < 0) lineEnd = source.Length;

            var line = source[lineStart..lineEnd].Trim();
            if (line.Contains("Aspose.Pdf", StringComparison.Ordinal)) continue;
            if (importsPdf && !importsWords) continue;

            offenders.Add($"{label}: {line}");
        }

        return offenders;
    }

    [Fact]
    public void EveryWordLoadFromAPath_ShouldGoThroughTheGuardedLoader()
    {
        var root = RepositoryRoot();
        var offenders = new List<string>();
        var productionDirectories = new[] { "Core", "Handlers", "Helpers", "Tools", "Errors" };
        var filesScanned = 0;
        var guardedSites = 0;

        foreach (var directory in productionDirectories)
        {
            var full = Path.Combine(root.FullName, directory);
            if (!Directory.Exists(full)) continue;

            foreach (var file in Directory.GetFiles(full, "*.cs", SearchOption.AllDirectories))
            {
                if (Exempt.ContainsKey(RepositoryRelative(file))) continue;

                var source = File.ReadAllText(file, Encoding.UTF8);
                filesScanned++;
                guardedSites += GuardedLoad.Matches(source).Count;

                offenders.AddRange(
                    UnguardedLoads(source, Path.GetRelativePath(root.FullName, file)));
            }
        }

        // Without these the guard reports success when it scanned nothing, or when it is pointed at
        // a tree that no longer contains the loads it is meant to police.
        Assert.True(filesScanned > 0, "No production source was scanned; the directory list is stale.");
        Assert.True(guardedSites > 0,
            "No call to GuardedWordLoader.Load was found, so either the loader was renamed or this "
            + "guard is looking at the wrong tree.");

        Assert.True(offenders.Count == 0,
            $"{offenders.Count} Word document(s) are loaded from a path without the resource guard. "
            + "Use GuardedWordLoader.Load so the external-resource policy applies to every entry "
            + $"point:\n  {string.Join("\n  ", offenders)}");
    }

    [Theory]
    [InlineData("new Document(path)")]
    [InlineData("new Aspose.Words.Document(path)")]
    [InlineData("new Aspose . Words . Document(path)")]
    [InlineData("new WordDocument(path)")]
    public void ALoadWrittenAnyOfTheWaysTheTypeCanBeSpelled_ShouldBeFlagged(string construction)
    {
        // Only the unqualified spelling was recognised, so moving a load to its fully qualified or
        // aliased name took it out of the guard's sight while every assertion stayed green
        // (§18.5.2).
        var source = "using Aspose.Words;\n"
                     + "public void Run(string path)\n"
                     + "{\n"
                     + $"    var doc = {construction};\n"
                     + "}\n";

        var offenders = UnguardedLoads(source, "synthetic");

        Assert.Single(offenders);
        Assert.Contains(construction, offenders[0], StringComparison.Ordinal);
    }

    [Fact]
    public void AnEmptyDocument_ShouldNotBeTreatedAsALoad()
    {
        const string source = "using Aspose.Words;\n"
                              + "public void Run()\n"
                              + "{\n"
                              + "    var doc = new Document();\n"
                              + "}\n";

        Assert.Empty(UnguardedLoads(source, "synthetic"));
    }

    [Fact]
    public void GoingThroughTheGuardedLoader_ShouldNotBeFlagged()
    {
        const string source = "using Aspose.Words;\n"
                              + "public void Run(string path)\n"
                              + "{\n"
                              + "    var doc = GuardedWordLoader.Load(path, bases);\n"
                              + "}\n";

        Assert.Empty(UnguardedLoads(source, "synthetic"));
    }

    [Fact]
    public void APdfDocument_ShouldNotBeMistakenForAWordOne()
    {
        const string source = "using Aspose.Pdf;\n"
                              + "public void Run(string path)\n"
                              + "{\n"
                              + "    var doc = new Document(path);\n"
                              + "}\n";

        Assert.Empty(UnguardedLoads(source, "synthetic"));
    }

    [Fact]
    public void EveryWordDocumentConstructedFromAPath_ShouldBeInTheGuardedLoader_BySymbol()
    {
        // R21-TST03. The regex above recognises a handful of spellings and guesses the namespace
        // from the file's imports, so `new Aspose.Words.Document(path)` in a file that also uses
        // Aspose.Pdf, or `using W = Aspose.Words.Document;`, walked past it. The semantic model
        // does not guess: it says which constructor was bound.
        var (compilation, unresolved) = FanOutBudgetFlowTests.ProductionCompilation();
        Assert.True(unresolved.Count == 0, "anchor symbols did not resolve: " + string.Join(", ", unresolved));

        var wordDocument = compilation.GetTypeByMetadataName("Aspose.Words.Document");
        Assert.NotNull(wordDocument);

        var offenders = new List<string>();
        var loadsSeen = 0;

        foreach (var tree in compilation.SyntaxTrees)
        {
            var file = RepositoryRelative(tree.FilePath);
            var model = compilation.GetSemanticModel(tree);

            foreach (var creation in tree.GetRoot().DescendantNodes()
                         .OfType<ObjectCreationExpressionSyntax>())
            {
                if (model.GetSymbolInfo(creation).Symbol is not IMethodSymbol ctor) continue;
                if (!SymbolEqualityComparer.Default.Equals(ctor.ContainingType, wordDocument)) continue;
                // Every constructor that takes an input — a path, a stream, anything — not only
                // the ones whose first parameter is a string: a stream-based load reaches the
                // same parser with the same external-resource handling (R22-TST02). Only the
                // parameterless blank document is out of scope.
                if (ctor.Parameters.Length == 0) continue;

                loadsSeen++;
                if (Exempt.ContainsKey(file)) continue;

                var line = creation.GetLocation().GetLineSpan().StartLinePosition.Line + 1;
                offenders.Add($"{file}:{line}: {creation.ToString().Trim()}");
            }
        }

        Assert.True(loadsSeen > 0, "no Aspose.Words.Document construction with an input was found at all");
        Assert.True(offenders.Count == 0,
            "Word documents constructed from an input outside Helpers/Word/GuardedWordLoader.cs: "
            + string.Join(" | ", offenders));
    }
}
