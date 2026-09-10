using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.Paragraph;

/// <summary>
///     R10-W01 (§22.5): a refused paragraph edit must leave the paragraph exactly as it was.
///     <para>
///         The handler applied font, then paragraph formatting, then line spacing, then style,
///         then tab stops, and only then the text — with the field check inside that last step.
///         So a refusal from any later step left every earlier one applied. Measured: an edit
///         refused for a field had already changed the paragraph's alignment from Left to Center,
///         and <c>IsModified</c> was still <c>false</c> because it is set after all of them.
///     </para>
///     <para>
///         Not a field-specific problem: an unknown style name refused the same way, with the same
///         formatting left behind. The earlier fixtures only passed <c>text</c> and compared
///         <c>document.GetText()</c>, so formatting residue was invisible to them.
///     </para>
/// </summary>
public class ParagraphEditAtomicityTests : WordHandlerTestBase
{
    private readonly EditParagraphWordHandler _handler = new();

    /// <summary>A document whose IF field spans three paragraphs.</summary>
    /// <returns>The document.</returns>
    private static Document ADocumentWithAFieldSpanningThreeParagraphs()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.Writeln("before");
        var field = builder.InsertField("IF 1 = 1 \"yes\" \"no\"");
        builder.Writeln();
        builder.Writeln("after");

        var start = field.Start!.ParentParagraph;
        var middle = new WordParagraph(document);
        middle.AppendChild(new Run(document, "carried by the field"));
        start.ParentNode.InsertAfter(middle, start);

        var tail = new WordParagraph(document);
        middle.ParentNode.InsertAfter(tail, middle);
        tail.AppendChild(field.End!);

        return document;
    }

    /// <summary>The paragraphs of a document's first body.</summary>
    /// <param name="document">The document.</param>
    /// <returns>The paragraphs.</returns>
    private static List<WordParagraph> Paragraphs(Document document)
    {
        return document.FirstSection.Body.Paragraphs.Cast<WordParagraph>().ToList();
    }

    /// <summary>Everything about a paragraph that an edit could change.</summary>
    /// <param name="paragraph">The paragraph to describe.</param>
    /// <returns>A comparable description.</returns>
    private static string Snapshot(WordParagraph paragraph)
    {
        var format = paragraph.ParagraphFormat;
        return string.Join("|",
            paragraph.GetText(),
            format.Alignment,
            format.StyleName,
            format.LineSpacingRule,
            format.LineSpacing.ToString("F2"),
            format.LeftIndent.ToString("F2"),
            format.SpaceBefore.ToString("F2"),
            format.TabStops.Count.ToString());
    }

    [Fact]
    public void AnEditRefusedForAField_ShouldLeaveTheFormattingAlone()
    {
        var document = ADocumentWithAFieldSpanningThreeParagraphs();
        var extents = FieldBoundaryHelper.FieldExtents.Of(document);
        var paragraphs = Paragraphs(document);

        var inside = paragraphs
            .Where(paragraph => extents.EnclosingField(paragraph) != null
                                || extents.WouldSplitAField(paragraph))
            .ToList();
        Assert.NotEmpty(inside);

        var target = inside[0];
        var index = paragraphs.IndexOf(target);
        var before = Snapshot(target);
        var context = CreateContext(document);

        // alignment is applied several steps before the field check that refuses this.
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["paragraphIndex"] = index,
            ["text"] = "replacement",
            ["alignment"] = "center"
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Equal(before, Snapshot(target));
        Assert.False(context.IsModified);
    }

    [Fact]
    public void AnEditRefusedForAnUnknownStyle_ShouldLeaveTheFormattingAlone()
    {
        // The same defect without a field anywhere: validation interleaved with mutation.
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("ordinary paragraph");

        var target = Paragraphs(document)[0];
        var before = Snapshot(target);
        var context = CreateContext(document);

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["paragraphIndex"] = 0,
            ["alignment"] = "center",
            ["styleName"] = "NoSuchStyleExistsHere"
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Equal(before, Snapshot(target));
        Assert.False(context.IsModified);
    }

    [Fact]
    public void AnUnrecognisedAlignment_ShouldBeRefusedAndLeaveTheParagraphAlone()
    {
        // This was a characterization test: the helper mapped anything it did not recognise to
        // Left, so a caller's typo silently left-aligned the paragraph. That is now a refusal,
        // and like every other refusal here it must leave the paragraph untouched (§23.13.1).
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("ordinary paragraph");

        var target = Paragraphs(document)[0];
        target.ParagraphFormat.Alignment = ParagraphAlignment.Right;

        var before = Snapshot(target);
        var context = CreateContext(document);

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["paragraphIndex"] = 0,
            ["alignment"] = "diagonally"
        });

        var refusal = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("Unknown alignment", refusal.Message, StringComparison.Ordinal);
        Assert.Equal(before, Snapshot(target));
        Assert.Equal(ParagraphAlignment.Right, target.ParagraphFormat.Alignment);
        Assert.False(context.IsModified);
    }

    [Fact]
    public void AnEditRefusedAfterTabStopsWereGiven_ShouldLeaveTheExistingTabStopsAlone()
    {
        // ApplyTabStops clears the paragraph's existing stops before adding the new ones, and it
        // runs before the text step. A refusal from that later step used to leave the paragraph
        // with the new stops and no way back.
        var document = ADocumentWithAFieldSpanningThreeParagraphs();
        var extents = FieldBoundaryHelper.FieldExtents.Of(document);
        var paragraphs = Paragraphs(document);

        var inside = paragraphs
            .Where(paragraph => extents.EnclosingField(paragraph) != null
                                || extents.WouldSplitAField(paragraph))
            .ToList();
        Assert.NotEmpty(inside);

        var target = inside[0];
        target.ParagraphFormat.TabStops.Add(new TabStop(72.0));

        var before = Snapshot(target);
        var context = CreateContext(document);

        var stops = new JsonArray
        {
            new JsonObject { ["position"] = 144.0, ["alignment"] = "center" }
        };

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["paragraphIndex"] = paragraphs.IndexOf(target),
            ["text"] = "replacement",
            ["tabStops"] = stops
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Equal(before, Snapshot(target));
        Assert.Equal(1, target.ParagraphFormat.TabStops.Count);
        Assert.Equal(72.0, target.ParagraphFormat.TabStops[0].Position);
        Assert.False(context.IsModified);
    }

    [Fact]
    public void AMalformedTabStopPosition_ShouldLeaveTheParagraphEntirelyAlone()
    {
        // §23.4:  was the one tab-stop field the preflight did not read, and
        // ApplyTabStops clears the existing stops before reading it. So a string there threw
        // after the old stops were gone, and after alignment had already been applied.
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("ordinary paragraph");

        var target = Paragraphs(document)[0];
        target.ParagraphFormat.TabStops.Add(new TabStop(72.0));

        var before = Snapshot(target);
        var context = CreateContext(document);

        var stops = new JsonArray
        {
            new JsonObject { ["position"] = "not-a-number", ["alignment"] = "center" }
        };

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["paragraphIndex"] = 0,
            ["alignment"] = "center",
            ["tabStops"] = stops
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Equal(before, Snapshot(target));
        Assert.Equal(1, target.ParagraphFormat.TabStops.Count);
        Assert.Equal(72.0, target.ParagraphFormat.TabStops[0].Position);
        Assert.False(context.IsModified);
    }

    [Fact]
    public void ValidTabStops_ShouldStillReplaceTheExistingOnes()
    {
        // The control: refusing a malformed entry must not stop a well-formed array from working.
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("ordinary paragraph");

        var target = Paragraphs(document)[0];
        target.ParagraphFormat.TabStops.Add(new TabStop(72.0));

        var stops = new JsonArray
        {
            new JsonObject { ["position"] = 144.0, ["alignment"] = "center" },
            new JsonObject { ["position"] = 216.0, ["alignment"] = "right" }
        };

        _handler.Execute(CreateContext(document), CreateParameters(new Dictionary<string, object?>
        {
            ["paragraphIndex"] = 0,
            ["tabStops"] = stops
        }));

        Assert.Equal(2, target.ParagraphFormat.TabStops.Count);
        Assert.Equal(144.0, target.ParagraphFormat.TabStops[0].Position);
    }

    [Fact]
    public void AnEditThatIsAccepted_ShouldStillApplyEverything()
    {
        // The control: preflight must not refuse a request that is entirely valid, and every part
        // of it must still take effect.
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("ordinary paragraph");

        var target = Paragraphs(document)[0];
        var context = CreateContext(document);

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["paragraphIndex"] = 0,
            ["text"] = "replacement",
            ["alignment"] = "center",
            ["spaceBefore"] = 12.0
        });

        _handler.Execute(context, parameters);

        Assert.Equal(ParagraphAlignment.Center, target.ParagraphFormat.Alignment);
        Assert.Equal(12.0, target.ParagraphFormat.SpaceBefore);
        Assert.Contains("replacement", target.GetText(), StringComparison.Ordinal);
        Assert.True(context.IsModified);
    }
}
