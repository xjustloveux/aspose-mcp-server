using System.Text.RegularExpressions;

namespace AsposeMcpServer.Tests.Infrastructure;

/// <summary>
///     R9-C01: every handler that writes into a directory the caller names must bound how much it
///     writes there, and every one of them must say which bound it uses.
///     <para>
///         A handler given an output <em>file</em> writes what the caller asked for, and its size
///         is the caller's business. A handler given an output <em>directory</em> decides for
///         itself how many files to put in it and how large they are, so it needs a ceiling — and
///         the conversion guard next door cannot speak for these, because it reads two named
///         source files (R9-C01).
///     </para>
///     <para>
///         The bounds in this codebase were not written to one pattern. Splitting counts work
///         units, extract-all sums bytes against a server setting, rendering prices pixels, and
///         conversion publishes through a bounded batch. All four are real ceilings; a scan that
///         recognised only one of them would have reported the other three as unbounded. What this
///         guard requires is that each handler uses one of them, named here, and that a handler
///         claiming to write a single file really does write one.
///     </para>
/// </summary>
public class FanOutBudgetInventoryTests : TestBase
{
    /// <summary>The ways a handler in this codebase bounds what it writes.</summary>
    private static readonly string[] KnownBudgets =
    [
        "RenderBudget", // total pixels, output files, output bytes
        "BoundedFileBatch", // one byte budget shared across a fan-out, published at once
        "BoundedFilePublisher", // per-file byte cap with transactional publish
        "MaxTotalWorkUnits", // splitting: source units x output files
        "MaxExtractAllBytes", // extract-all: cumulative bytes against a server setting
        "PixelBudget" // rendering: pixels before the raster is attempted
    ];

    /// <summary>
    ///     Handlers that take an output directory but write exactly one file into it. Named
    ///     individually, because "writes one file" is a claim about behaviour that a scan cannot
    ///     make on its own — and the claim is checked below.
    /// </summary>
    private static readonly string[] SingleFileWriters =
    [
        "Handlers/BarCode/Generate/GenerateBarcodeHandler.cs",
        "Handlers/Excel/OleObject/ExtractExcelOleObjectHandler.cs",
        "Handlers/Ocr/Preprocessing/OcrPreprocessingHandlerBase.cs",
        "Handlers/PowerPoint/OleObject/ExtractPptOleObjectHandler.cs",
        "Handlers/Word/File/CreateFromTemplateWordHandler.cs",
        "Handlers/Word/File/CreateWordDocumentHandler.cs",
        "Handlers/Word/File/MergeWordDocumentsHandler.cs",
        "Handlers/Word/OleObject/ExtractWordOleObjectHandler.cs"
    ];

    /// <summary>Anything that puts bytes on disk at a path.</summary>
    private static readonly Regex DiskSink = new(
        @"\.Save\(\s*(?!.*Stream)|\.ToImage\(|File\.WriteAllBytes\(|File\.WriteAllText\(|"
        + @"File\.WriteAllLines\(|new FileStream\(|File\.Create\(|File\.Copy\(|\.ExtractToFile\(",
        RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>
    ///     A handler that writes into a directory the caller chose. Two spellings were recognised
    ///     before, so a handler naming the same parameter anything else fell out of the scan
    ///     entirely (§21.6).
    /// </summary>
    private static readonly Regex CallerNamedDirectory = new(
        @"outputDirectory|outputDir\b|outputFolder|targetDirectory|targetDir\b|destinationDirectory|"
        + @"extractDirectory|extractTo\b|outputPath.*Directory\.CreateDirectory",
        RegexOptions.Compiled, TimeSpan.FromSeconds(5));

    /// <summary>
    ///     Strips comment text, so a budget named only in a comment does not read as one the code
    ///     applies. This is still a textual check and not a control-flow proof — what it can show
    ///     is that the name is used, not that it dominates the sink (§21.6, recorded as residual).
    /// </summary>
    /// <param name="text">The source text.</param>
    /// <returns>The text with line and block comments removed.</returns>
    private static string WithoutComments(string text)
    {
        var stripped = Regex.Replace(text, @"/\*.*?\*/", " ",
            RegexOptions.Singleline, TimeSpan.FromSeconds(5));
        return Regex.Replace(stripped, @"//[^\n]*", " ", RegexOptions.None, TimeSpan.FromSeconds(5));
    }

    /// <summary>Walks up from the test binaries to the repository root.</summary>
    /// <returns>The repository root directory.</returns>
    private static DirectoryInfo RepositoryRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "AsposeMcpServer.csproj")))
            dir = dir.Parent;

        Assert.NotNull(dir);
        return dir;
    }

    /// <summary>Every handler that writes into a caller-named directory, with its source.</summary>
    /// <returns>Repository-relative path and content for each such handler.</returns>
    private static List<(string Path, string Text)> HandlersWritingIntoACallerNamedDirectory()
    {
        var root = RepositoryRoot();
        var handlers = Path.Combine(root.FullName, "Handlers");

        return Directory.EnumerateFiles(handlers, "*.cs", SearchOption.AllDirectories)
            .Select(path => (Path: Path.GetRelativePath(root.FullName, path).Replace('\\', '/'),
                Text: File.ReadAllText(path)))
            .Where(file => DiskSink.IsMatch(file.Text) && CallerNamedDirectory.IsMatch(file.Text))
            .OrderBy(file => file.Path, StringComparer.Ordinal)
            .ToList();
    }

    [Fact]
    public void EveryHandlerWritingIntoACallerNamedDirectory_ShouldBoundWhatItWrites()
    {
        var unbounded = HandlersWritingIntoACallerNamedDirectory()
            .Where(file => !SingleFileWriters.Contains(file.Path, StringComparer.Ordinal))
            .Where(file => !KnownBudgets.Any(budget =>
                WithoutComments(file.Text).Contains(budget, StringComparison.Ordinal)))
            .Select(file => file.Path)
            .ToList();

        Assert.True(unbounded.Count == 0,
            "these handlers write into a directory the caller names with no ceiling on how much: "
            + string.Join(", ", unbounded)
            + ". Give it one of " + string.Join("/", KnownBudgets)
            + ", or add it to SingleFileWriters if it really writes one file.");
    }

    [Fact]
    public void EverySingleFileWriter_ShouldStillWriteJustOne()
    {
        // The claim that keeps a handler out of the check above, checked: a sink inside a loop is
        // a handler that writes as many files as the loop runs, whatever the list says.
        var contradicted = new List<string>();

        foreach (var (path, text) in HandlersWritingIntoACallerNamedDirectory()
                     .Where(file => SingleFileWriters.Contains(file.Path, StringComparer.Ordinal)))
        {
            var lines = text.Split('\n');
            var depth = 0;
            var loopDepths = new List<int>();

            foreach (var line in lines)
            {
                if (Regex.IsMatch(line, @"^\s*(for|foreach|while)\s*\(", RegexOptions.None,
                        TimeSpan.FromSeconds(5)))
                    loopDepths.Add(depth);

                if (loopDepths.Count > 0 && DiskSink.IsMatch(line))
                {
                    contradicted.Add($"{path}: writes inside a loop");
                    break;
                }

                depth += line.Count(c => c == '{') - line.Count(c => c == '}');
                loopDepths.RemoveAll(d => d >= depth);
            }
        }

        Assert.True(contradicted.Count == 0,
            "these are listed as writing one file but write inside a loop: "
            + string.Join(", ", contradicted));
    }

    [Fact]
    public void TheSingleFileWriterList_ShouldNotHoldNamesThatNoLongerApply()
    {
        // A list that outlives what it describes stops being an inventory and becomes an excuse.
        var actual = HandlersWritingIntoACallerNamedDirectory()
            .Select(file => file.Path)
            .ToHashSet(StringComparer.Ordinal);

        var stale = SingleFileWriters.Where(name => !actual.Contains(name)).ToList();

        Assert.True(stale.Count == 0,
            "these are listed as single-file writers but no longer write into a caller-named "
            + "directory: " + string.Join(", ", stale));
    }

    [Fact]
    public void TheScan_ShouldFindTheHandlersItIsMeantToCover()
    {
        // Without this, a rename of the parameter would empty the scan and leave every assertion
        // above vacuously true.
        var found = HandlersWritingIntoACallerNamedDirectory().Select(file => file.Path).ToList();

        Assert.Contains("Handlers/Excel/FileOperations/SplitWorkbookHandler.cs", found);
        Assert.Contains("Handlers/PowerPoint/FileOperations/SplitPresentationHandler.cs", found);
        Assert.Contains("Handlers/Word/OleObject/ExtractAllWordOleObjectHandler.cs", found);
        Assert.True(found.Count >= 15, $"only {found.Count} handlers matched; the scan looks broken");
    }

    [Fact]
    public void ABudgetNamedOnlyInAComment_ShouldNotCountAsOne()
    {
        // The shape the previous guard accepted: the word appears in the file, so the file read as
        // bounded. Stripping comments is what makes the check about the code (§21.6).
        const string commentOnly = "class X { // RenderBudget applies here\n void Go() { } }";
        const string realUse = "class X { void Go() { RenderBudget.EnsureOutputCount(1); } }";

        Assert.DoesNotContain("RenderBudget", WithoutComments(commentOnly), StringComparison.Ordinal);
        Assert.Contains("RenderBudget", WithoutComments(realUse), StringComparison.Ordinal);
    }

    [Fact]
    public void ADirectoryParameterUnderAnotherName_ShouldStillBeScanned()
    {
        // Recognising only two spellings meant a handler could leave the scan by renaming its
        // parameter, which is not a change to what it writes.
        foreach (var spelling in new[]
                 {
                     "outputDirectory", "outputDir", "outputFolder", "targetDirectory",
                     "targetDir", "destinationDirectory", "extractDirectory"
                 })
            Assert.True(CallerNamedDirectory.IsMatch($"var {spelling} = p.Path;"),
                $"a handler naming its directory '{spelling}' is invisible to this scan");
    }

    [Fact]
    public void TheDiskSinkPattern_ShouldRecogniseTheSpellingsInUse()
    {
        // A sink pattern that matches nothing makes every guard above vacuously green.
        foreach (var spelling in new[]
                 {
                     "document.Save(path);", "sheet.ToImage(0, path);",
                     "File.WriteAllBytes(path, bytes);", "File.WriteAllText(path, text);",
                     "new FileStream(path, FileMode.Create)", "entry.ExtractToFile(path);"
                 })
            Assert.True(DiskSink.IsMatch(spelling),
                $"the disk-sink pattern does not recognise '{spelling}'");
    }
}
