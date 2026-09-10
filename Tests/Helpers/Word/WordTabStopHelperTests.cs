using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers.Word;

/// <summary>
///     R13-W01: the one reader for the <c>tabStops</c> array, and the contract it holds callers to.
///     <para>
///         There were five readers before this, and they disagreed on every question a caller could
///         ask. A null entry was skipped by one and became a stop at position 0 in the others; a
///         missing <c>position</c> meant 0 everywhere, which is the left margin and so no stop at
///         all; a misspelt alignment was silently taken as left by four of them, and
///         <c>add_styled</c> matched PascalCase only, so <c>"center"</c> worked in three
///         operations and quietly did not in the fourth.
///     </para>
///     <para>
///         The caps here are measured against Aspose.Words 23.10, not chosen: 1,500 stops on one
///         paragraph survive a save and reload unchanged while 2,000 come back as 1,947, and a
///         position of 1e9 is accepted and reads back as 107,374,182.35 where the twip field
///         saturates. Both are cases where the request would be accepted and then silently not be
///         what was stored.
///     </para>
/// </summary>
public class WordTabStopHelperTests : TestBase
{
    /// <summary>Where the one mapping from a tab-stop name to its enum is allowed to live.</summary>
    private static readonly string[] MayMapTabNames =
    [
        "Helpers/Word/WordParagraphHelper.cs"
    ];

    /// <summary>An array of the given entries.</summary>
    /// <param name="entries">The entries.</param>
    /// <returns>The array.</returns>
    private static JsonArray Array(params JsonNode?[] entries)
    {
        return new JsonArray(entries);
    }

    /// <summary>An array of the given number of well-formed stops.</summary>
    /// <param name="count">How many.</param>
    /// <returns>The array.</returns>
    private static JsonArray Stops(int count)
    {
        var array = new JsonArray();
        for (var i = 1; i <= count; i++) array.Add(new JsonObject { ["position"] = (double)i });
        return array;
    }

    [Fact]
    public void NoOtherFile_ShouldMapTabStopNamesForItself()
    {
        // Each of the five readers had its own switch, and they did not agree: four fell back to
        // left or none for anything unrecognised, one matched PascalCase only, and two omitted
        // half the vocabulary. Finding the existing mapping has to be easier than writing another.
        var root = RepositoryRoot();

        var offenders = Directory
            .EnumerateFiles(root.FullName, "*.cs", SearchOption.AllDirectories)
            .Where(IsProductionSource)
            .Select(path => (Path: Path.GetRelativePath(root.FullName, path).Replace('\\', '/'),
                Text: File.ReadAllText(path)))
            .Where(file => !MayMapTabNames.Contains(file.Path, StringComparer.Ordinal))
            .Where(file => file.Text.Contains("=> TabAlignment.", StringComparison.Ordinal)
                           || file.Text.Contains("=> TabLeader.", StringComparison.Ordinal))
            .Select(file => file.Path)
            .ToList();

        Assert.True(offenders.Count == 0,
            "These files map tab alignment or leader names to their enums for themselves, instead "
            + "of using WordParagraphHelper through WordTabStopHelper, which is how five handlers "
            + "came to disagree about what the same request meant (R13-W01): "
            + string.Join(", ", offenders));
    }

    /// <summary>Whether a path is production source rather than a test or a build artefact.</summary>
    /// <param name="path">The file to judge.</param>
    /// <returns><c>true</c> when it is part of the server.</returns>
    private static bool IsProductionSource(string path)
    {
        var separator = Path.DirectorySeparatorChar;
        return !path.Contains($"{separator}Tests{separator}", StringComparison.Ordinal)
               && !path.Contains($"{separator}obj{separator}", StringComparison.Ordinal)
               && !path.Contains($"{separator}bin{separator}", StringComparison.Ordinal);
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

    [Fact]
    public void NoArrayAtAll_ShouldResolveToNothingRatherThanRefuse()
    {
        // Callers pass it through whether or not the request named any, and "none given" means
        // "leave the paragraph's stops alone", not "this request is malformed".
        Assert.Empty(WordTabStopHelper.Resolve(null));
        Assert.Empty(WordTabStopHelper.Resolve([]));
    }

    [Fact]
    public void AnEntryThatIsNotAnObject_ShouldBeRefusedRatherThanBecomeAStopAtZero()
    {
        // A JSON null in the array used to produce a tab stop on the left margin, which is a stop
        // that does nothing, reported as success.
        var refusal = Assert.Throws<ArgumentException>(() =>
            WordTabStopHelper.Resolve(Array(new JsonObject { ["position"] = 72.0 }, null)));

        Assert.Contains("Tab stop 1", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnEntryThatIsABareNumber_ShouldBeRefused()
    {
        var refusal = Assert.Throws<ArgumentException>(() =>
            WordTabStopHelper.Resolve(Array(JsonValue.Create(72.0))));

        Assert.Contains("expected an object", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AMissingPosition_ShouldBeRefused()
    {
        var refusal = Assert.Throws<ArgumentException>(() =>
            WordTabStopHelper.Resolve(Array(new JsonObject { ["alignment"] = "center" })));

        Assert.Contains("'position' is required", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void APositionThatIsNotANumber_ShouldBeRefusedBeforeAnythingIsApplied()
    {
        var refusal = Assert.Throws<ArgumentException>(() =>
            WordTabStopHelper.Resolve(Array(new JsonObject { ["position"] = "72pt" })));

        Assert.Contains("must be a number", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void APositionThatIsNotFinite_ShouldBeRefused()
    {
        var refusal = Assert.Throws<ArgumentException>(() =>
            WordTabStopHelper.Resolve(Array(new JsonObject
            {
                ["position"] = JsonValue.Create(double.NaN)
            })));

        Assert.Contains("finite", refusal.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(-1584.0)]
    [InlineData(0.0)]
    [InlineData(1584.0)]
    public void APositionOnAPageWordWillLayOut_ShouldBeAccepted(double position)
    {
        var resolved = WordTabStopHelper.Resolve(Array(new JsonObject { ["position"] = position }));

        Assert.Equal(position, Assert.Single(resolved).Position);
    }

    [Theory]
    [InlineData(1584.5)]
    [InlineData(-1584.5)]
    [InlineData(1e9)]
    public void APositionOffAnyPage_ShouldBeRefusedRatherThanSilentlyRewritten(double position)
    {
        // 1e9 is not rejected further down: it is accepted and reads back as 107,374,182.35, the
        // point at which the underlying twip field saturates.
        var refusal = Assert.Throws<ArgumentException>(() =>
            WordTabStopHelper.Resolve(Array(new JsonObject { ["position"] = position })));

        Assert.Contains("1584", refusal.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("center", TabAlignment.Center)]
    [InlineData("Center", TabAlignment.Center)]
    [InlineData("CENTER", TabAlignment.Center)]
    [InlineData("decimal", TabAlignment.Decimal)]
    [InlineData("bar", TabAlignment.Bar)]
    [InlineData("clear", TabAlignment.Clear)]
    public void AnAlignmentName_ShouldBeReadWithoutRegardToCase(string name, TabAlignment expected)
    {
        // add_styled matched PascalCase only, so "center" was honoured by the header, footer and
        // edit operations and silently read as left by that one.
        var resolved = WordTabStopHelper.Resolve(Array(new JsonObject
        {
            ["position"] = 72.0,
            ["alignment"] = name
        }));

        Assert.Equal(expected, Assert.Single(resolved).Alignment);
    }

    [Theory]
    [InlineData("dots", TabLeader.Dots)]
    [InlineData("Dots", TabLeader.Dots)]
    [InlineData("middledot", TabLeader.MiddleDot)]
    [InlineData("MiddleDot", TabLeader.MiddleDot)]
    [InlineData("heavy", TabLeader.Heavy)]
    public void ALeaderName_ShouldBeReadWithoutRegardToCase(string name, TabLeader expected)
    {
        var resolved = WordTabStopHelper.Resolve(Array(new JsonObject
        {
            ["position"] = 72.0,
            ["leader"] = name
        }));

        Assert.Equal(expected, Assert.Single(resolved).Leader);
    }

    [Fact]
    public void AMisspeltAlignment_ShouldBeRefusedRatherThanTakenAsLeft()
    {
        var refusal = Assert.Throws<ArgumentException>(() =>
            WordTabStopHelper.Resolve(Array(new JsonObject
            {
                ["position"] = 72.0,
                ["alignment"] = "centre"
            })));

        Assert.Contains("centre", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnAlignmentThatIsNotAString_ShouldSayThatRatherThanThrowFromInsideJson()
    {
        var refusal = Assert.Throws<ArgumentException>(() =>
            WordTabStopHelper.Resolve(Array(new JsonObject
            {
                ["position"] = 72.0,
                ["alignment"] = 3
            })));

        Assert.Contains("must be a string", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AsManyStopsAsAParagraphKeeps_ShouldBeAccepted()
    {
        Assert.Equal(WordTabStopHelper.MaxTabStops,
            WordTabStopHelper.Resolve(Stops(WordTabStopHelper.MaxTabStops)).Count);
    }

    [Fact]
    public void MoreStopsThanAParagraphKeeps_ShouldBeRefused()
    {
        // Above the bound the document that comes back is not the one that was asked for, and
        // nothing between here and the file says so.
        var refusal = Assert.Throws<ArgumentException>(() =>
            WordTabStopHelper.Resolve(Stops(WordTabStopHelper.MaxTabStops + 1)));

        Assert.Contains("more than the", refusal.Message, StringComparison.Ordinal);
    }

    [SkippableFact]
    public void TheCountCap_ShouldBeAFigureThatSurvivesASaveAndReload()
    {
        // The cap is a measurement, and a measurement that is never repeated is a comment. This
        // repeats it: a paragraph given MaxTabStops stops still has them after a round trip.
        SkipInEvaluationMode(AsposeLibraryType.Words, "Evaluation mode truncates the document");

        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("body");

        var paragraph = document.FirstSection.Body.FirstParagraph;
        foreach (var stop in WordTabStopHelper.Resolve(Stops(WordTabStopHelper.MaxTabStops)))
            paragraph.ParagraphFormat.TabStops.Add(stop);

        var path = CreateTestFilePath("tab_stop_cap.docx");
        document.Save(path);

        var reloaded = new Document(path);
        Assert.Equal(WordTabStopHelper.MaxTabStops,
            reloaded.FirstSection.Body.FirstParagraph.ParagraphFormat.TabStops.Count);
    }
}
