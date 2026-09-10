namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     Every production file that takes hold of an <c>Aspose.Slides.Presentation</c> must take the
///     gate that serialises this process's use of the library — and must take it once per place
///     that does.
///     <para>
///         Two earlier versions of this guard were weaker than they read. The first looked for
///         <c>new Presentation(</c>, so a constructor stood in for the boundary and every path
///         working on a presentation it was <em>handed</em> was invisible (R9-S01). The second
///         asked only whether the file contained <c>SlidesGate.Enter()</c> anywhere, so in a file
///         with two gated operations, removing one left the other covering for it and the guard
///         stayed green (§21.6).
///     </para>
///     <para>
///         What is checked now is an inventory: each file, how many places in it enter the
///         library, and how many gates it takes. Removing a gate makes the counts disagree, and so
///         does adding an acquisition — either kind of drift has to be reconciled deliberately,
///         with the reason written down.
///     </para>
///     <para>
///         <b>What this cannot prove.</b> It counts tokens in a file. It does not know whether a
///         gate and an acquisition are in the same method, whether the gate is taken before the
///         acquisition, or whether it is held for the acquisition's whole life — a gate moved to
///         an unrelated method keeps the count and passes. Proving that needs syntax and
///         control-flow analysis (Roslyn), which this does not do; the runtime fixtures in
///         <c>SlidesGateSessionCoverageTests</c> are what actually demonstrate the gate holding,
///         one operation at a time. This guard's job is narrower: to notice that the set of places
///         entering the library has changed (§23.8).
///     </para>
/// </summary>
public class SlidesGateCoverageTests : TestBase
{
    /// <summary>The production directories that may contain Aspose.Slides work.</summary>
    private static readonly string[] ProductionRoots = ["Core", "Handlers", "Helpers", "Tools"];

    /// <summary>
    ///     Every way production code comes to hold a presentation: constructing one, taking one
    ///     from a session, or casting an untyped document to one.
    /// </summary>
    private static readonly string[] AcquisitionTokens =
    [
        "new Presentation(",
        "GetDocument<Presentation>",
        "(Presentation)",
        "is Aspose.Slides.Presentation",
        "is Presentation "
    ];

    /// <summary>
    ///     Each file that enters Aspose.Slides, how many gates it must take, and what they are
    ///     for. The count is the guard: a file with two gated operations that loses one no longer
    ///     matches.
    /// </summary>
    private static readonly (string File, int Gates, int Acquisitions, string Why)[] Inventory =
    [
        ("Core/Conversion/DocumentConverter.cs", 3, 2,
            "ConvertPowerPointDocument and ConvertPowerPointToStream — the file and in-memory "
            + "conversion boundaries, taken here rather than at each caller — plus SlideCountOf, "
            + "which reads the slide count before either of those has taken its hold"),
        ("Core/Session/DocumentContext.cs", 1, 4,
            "Create takes it before the presentation is loaded and holds it until the context is "
            + "disposed"),
        ("Core/Session/DocumentSession.cs", 1, 1,
            "Dispose releases the presentation, which re-enters the library on whichever thread "
            + "closed the session"),
        ("Core/Session/DocumentSessionManager.cs", 2, 2,
            "LoadPresentation, and SaveDocumentToFile — which covers explicit save, save on "
            + "close, all three disconnect behaviours and auto-save"),
        ("Tools/Conversion/ConvertDocumentTool.cs", 1, 2,
            "the file-mode branch constructs its own presentation; the session branch is covered "
            + "by DocumentConverter's own gate"),
        ("Handlers/PowerPoint/FileOperations/ConvertPresentationHandler.cs", 1, 2, "opens its own"),
        ("Handlers/PowerPoint/FileOperations/CreatePresentationHandler.cs", 1, 1, "builds its own"),
        ("Handlers/PowerPoint/FileOperations/MergePresentationsHandler.cs", 1, 2,
            "opens the master and each source; one hold covers the whole operation"),
        ("Handlers/PowerPoint/FileOperations/SplitPresentationHandler.cs", 1, 3,
            "opens the source and builds each output; one hold covers the whole operation"),
        ("Handlers/PowerPoint/Layout/ApplyThemeHandler.cs", 1, 1, "opens the theme presentation")
    ];

    /// <summary>
    ///     Walks up from the test binaries to the repository root so the scan works from any
    ///     working directory the runner chooses.
    /// </summary>
    /// <returns>The repository root directory.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir;
    }

    /// <summary>How many times a string occurs in another.</summary>
    /// <param name="text">The text to search.</param>
    /// <param name="token">What to count.</param>
    /// <returns>The number of occurrences.</returns>
    private static int Count(string text, string token)
    {
        var found = 0;
        for (var at = text.IndexOf(token, StringComparison.Ordinal);
             at >= 0;
             at = text.IndexOf(token, at + token.Length, StringComparison.Ordinal))
            found++;

        return found;
    }

    /// <summary>
    ///     Every production source file that takes hold of a presentation, with its text.
    /// </summary>
    /// <returns>Repository-relative path and content for each such file.</returns>
    private static List<(string Path, string Text)> FilesHoldingAPresentation()
    {
        var root = RepositoryRoot();

        return ProductionRoots
            .Select(name => Path.Combine(root.FullName, name))
            .Where(Directory.Exists)
            .SelectMany(dir => Directory.EnumerateFiles(dir, "*.cs", SearchOption.AllDirectories))
            .Select(path => (Path: Path.GetRelativePath(root.FullName, path).Replace('\\', '/'),
                Text: File.ReadAllText(path)))
            .Where(file => AcquisitionTokens.Any(token =>
                file.Text.Contains(token, StringComparison.Ordinal)))
            .OrderBy(file => file.Path, StringComparer.Ordinal)
            .ToList();
    }

    [Fact]
    public void EachFileThatEntersTheLibrary_ShouldTakeExactlyTheGatesItsInventoryEntrySays()
    {
        var root = RepositoryRoot();
        var wrong = new List<string>();

        foreach (var (file, gates, expectedAcquisitions, why) in Inventory)
        {
            var path = Path.Combine(root.FullName, file);
            if (!File.Exists(path))
            {
                wrong.Add($"{file}: listed in the inventory but no longer exists");
                continue;
            }

            var text = File.ReadAllText(path);

            var found = Count(text, "SlidesGate.Enter()");
            if (found != gates)
                wrong.Add($"{file}: {found} gate(s), inventory says {gates} — {why}");

            // Counting gates alone catches a gate that was removed. It does not catch a place
            // that was *added*: a new acquisition in a listed file left the gate count unchanged
            // and passed (§22.9). Both numbers are recorded, so either kind of drift disagrees.
            var acquisitions = AcquisitionTokens.Sum(token => Count(text, token));
            if (acquisitions != expectedAcquisitions)
                wrong.Add($"{file}: {acquisitions} place(s) take hold of a presentation, inventory "
                          + $"says {expectedAcquisitions}. A new one needs its own gate and its own "
                          + "line here.");
        }

        Assert.True(wrong.Count == 0,
            "the Slides gate inventory no longer matches the code:\n  " + string.Join("\n  ", wrong));
    }

    [Fact]
    public void NoFileOutsideTheInventory_ShouldTakeHoldOfAPresentation()
    {
        // The other direction: a new handler that opens a presentation is a new place that must be
        // gated, and it has to be added here with a reason rather than inheriting silence.
        var listed = Inventory.Select(entry => entry.File).ToHashSet(StringComparer.Ordinal);

        var unlisted = FilesHoldingAPresentation()
            .Select(file => file.Path)
            .Where(path => !listed.Contains(path))
            .ToList();

        Assert.True(unlisted.Count == 0,
            "these files take hold of an Aspose.Slides.Presentation but are not in the gate "
            + "inventory: " + string.Join(", ", unlisted));
    }

    [Fact]
    public void TheScan_ShouldStillFindEveryFileTheInventoryLists()
    {
        // Without this, a change that stopped the scan matching anything would leave the guard
        // above vacuously green.
        var scanned = FilesHoldingAPresentation().Select(file => file.Path).ToHashSet(StringComparer.Ordinal);

        var missed = Inventory.Select(entry => entry.File)
            .Where(file => !scanned.Contains(file))
            .ToList();

        Assert.True(missed.Count == 0,
            "the acquisition scan no longer reaches these inventoried files, so it is not "
            + "measuring what it claims: " + string.Join(", ", missed));
        Assert.True(scanned.Count >= Inventory.Length);
    }

    [Fact]
    public void TheCountsAreTextual_AndTheGuardSaysSo()
    {
        // The mutations §23.8 names, demonstrated on strings rather than by editing the tree: a
        // gate in an unrelated method, a gate after the acquisition, and a gate inside a comment
        // or a string literal all leave the count unchanged. This test exists so the limitation
        // is written down as executable fact rather than only in prose — nothing above detects
        // any of these, and a reader should not have to discover that by trusting the name.
        const string gateInAnUnrelatedMethod =
            "void A() { var p = new Presentation(); } void B() { using var g = SlidesGate.Enter(); }";
        const string gateAfterTheAcquisition =
            "void A() { var p = new Presentation(); using var g = SlidesGate.Enter(); }";
        const string gateInAComment =
            "void A() { var p = new Presentation(); /* SlidesGate.Enter() */ }";

        foreach (var shape in new[] { gateInAnUnrelatedMethod, gateAfterTheAcquisition, gateInAComment })
        {
            Assert.Equal(1, Count(shape, "new Presentation("));
            Assert.Equal(1, Count(shape, "SlidesGate.Enter()"));
        }

        // All three have one acquisition and one gate, so all three satisfy the inventory: the
        // counts are identical to a correctly gated operation's. That is the boundary of what
        // counting can do, and SlidesGateSessionCoverageTests is what proves the gate holds.
        Assert.Equal(Count(gateInAComment, "new Presentation("),
            Count("void A() { using var g = SlidesGate.Enter(); var p = new Presentation(); }",
                "new Presentation("));
    }

    [Fact]
    public void TheGateItself_ShouldNotBeCountedAsAHolder()
    {
        // SlidesGate.cs names no presentation, so a scan that matched it would be matching the
        // word rather than the boundary.
        var scanned = FilesHoldingAPresentation().Select(file => file.Path).ToList();

        Assert.DoesNotContain(scanned,
            path => path.EndsWith("SlidesGate.cs", StringComparison.Ordinal));
    }
}
