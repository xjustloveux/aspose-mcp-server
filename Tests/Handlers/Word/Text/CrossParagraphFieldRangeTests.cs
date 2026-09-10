using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Handlers.Word.Paragraph;
using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Tests.Infrastructure;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

/// <summary>
///     R9-W02 (§21.3): the run-level field protection did not extend to whole paragraphs.
///     <para>
///         <c>FieldBoundaryHelper.FieldExtents</c> finds a field that spans several paragraphs, and
///         the run loops in these handlers use it. Three places did not: two removed whole
///         paragraphs in a loop with no check at all, and the third appended a run to a paragraph
///         whose every run was inside a field — so a range crossing a multi-paragraph field could
///         take one of its markers and leave the other orphaned, or write the caller's text into
///         a field's result.
///     </para>
///     <para>
///         Each test here asserts on the <em>document</em>, not on a file: a refused operation
///         never reaches a save, so file bytes would be unchanged either way and prove nothing.
///     </para>
/// </summary>
public class CrossParagraphFieldRangeTests : WordHandlerTestBase
{
    /// <summary>
    ///     Builds a document whose IF field starts in paragraph 1 and ends in paragraph 3, so
    ///     paragraph 2 lies wholly inside it.
    /// </summary>
    /// <returns>The document.</returns>
    private static Document ADocumentWithAFieldSpanningThreeParagraphs()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);

        builder.Writeln("before");
        var field = builder.InsertField("IF 1 = 1 \"yes\" \"no\"");
        builder.Writeln();
        builder.Writeln("after");

        // A field builder-inserted lives in one paragraph. Moving its end marker into a later one
        // is what makes it span several — the shape a sibling walk from the start marker can never
        // reach, and the shape these handlers mishandled.
        var start = field.Start!.ParentParagraph;
        var middle = new WordParagraph(document);
        middle.AppendChild(new Run(document, "carried by the field"));
        start.ParentNode.InsertAfter(middle, start);

        var tail = new WordParagraph(document);
        middle.ParentNode.InsertAfter(tail, middle);
        tail.AppendChild(field.End!);

        return document;
    }

    /// <summary>The paragraphs of a document's first body, as a list.</summary>
    /// <param name="document">The document.</param>
    /// <returns>The paragraphs.</returns>
    private static List<WordParagraph> Paragraphs(Document document)
    {
        return document.FirstSection.Body.Paragraphs.Cast<WordParagraph>().ToList();
    }

    /// <summary>Whether every field in the document still has both of its markers.</summary>
    /// <param name="document">The document to check.</param>
    /// <returns><c>true</c> when no field lost a marker.</returns>
    private static bool EveryFieldIsIntact(Document document)
    {
        return document.Range.Fields
            .All(field => field.Start != null && field.End != null);
    }

    [Fact]
    public void TheHelper_ShouldSeeAParagraphAFieldRunsThrough()
    {
        // The premise of every test below: if this is false the fixture does not build the shape
        // it claims and the rest would pass vacuously.
        var document = ADocumentWithAFieldSpanningThreeParagraphs();
        var extents = FieldBoundaryHelper.FieldExtents.Of(document);
        var paragraphs = Paragraphs(document);

        Assert.Contains(paragraphs, paragraph => extents.WouldSplitAField(paragraph));
    }

    [Fact]
    public void AFieldWhollyInsideOneParagraph_ShouldNotCountAsSplit()
    {
        // The control that keeps the guard from refusing every deletion: a field contained in one
        // paragraph goes with it, which is a complete removal and always was allowed.
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("before");
        builder.InsertField(FieldType.FieldDate, false);
        builder.Writeln();
        builder.Writeln("after");

        var extents = FieldBoundaryHelper.FieldExtents.Of(document);

        Assert.All(Paragraphs(document),
            paragraph => Assert.False(extents.WouldSplitAField(paragraph)));
    }

    [Fact]
    public void DeletingARangeThroughAField_ShouldBeRefusedAndChangeNothing()
    {
        var document = ADocumentWithAFieldSpanningThreeParagraphs();
        var before = document.GetText();
        var fieldCount = document.Range.Fields.Count;

        var handler = new DeleteRangeWordTextHandler();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["startParagraphIndex"] = 0,
            ["endParagraphIndex"] = Paragraphs(document).Count - 1,
            ["startCharIndex"] = 0,
            ["endCharIndex"] = 1
        });

        Assert.Throws<ArgumentException>(() => handler.Execute(CreateContext(document), parameters));

        Assert.Equal(before, document.GetText());
        Assert.Equal(fieldCount, document.Range.Fields.Count);
        Assert.True(EveryFieldIsIntact(document));
    }

    [Fact]
    public void DeletingTextThroughAField_ShouldBeRefusedAndChangeNothing()
    {
        var document = ADocumentWithAFieldSpanningThreeParagraphs();
        var before = document.GetText();
        var fieldCount = document.Range.Fields.Count;

        var handler = new DeleteWordTextHandler();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["startParagraphIndex"] = 0,
            ["endParagraphIndex"] = Paragraphs(document).Count - 1,
            ["startRunIndex"] = 0
        });

        Assert.Throws<ArgumentException>(() => handler.Execute(CreateContext(document), parameters));

        Assert.Equal(before, document.GetText());
        Assert.Equal(fieldCount, document.Range.Fields.Count);
        Assert.True(EveryFieldIsIntact(document));
    }

    [Fact]
    public void EditingAParagraphInsideAField_ShouldBeRefusedAndChangeNothing()
    {
        var document = ADocumentWithAFieldSpanningThreeParagraphs();
        var extents = FieldBoundaryHelper.FieldExtents.Of(document);
        var paragraphs = Paragraphs(document);

        var inside = paragraphs
            .Where(paragraph => extents.EnclosingField(paragraph) != null
                                || extents.WouldSplitAField(paragraph))
            .ToList();
        Assert.NotEmpty(inside);

        var index = paragraphs.IndexOf(inside[0]);

        var before = document.GetText();

        var handler = new EditParagraphWordHandler();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["paragraphIndex"] = index,
            ["text"] = "replacement"
        });

        Assert.Throws<ArgumentException>(() => handler.Execute(CreateContext(document), parameters));

        Assert.Equal(before, document.GetText());
        Assert.DoesNotContain("replacement", document.GetText(), StringComparison.Ordinal);
        Assert.True(EveryFieldIsIntact(document));
    }

    [Fact]
    public void AnOrdinaryRangeDeletion_ShouldStillWork()
    {
        // The control for the handlers: the guard must not refuse a range with no field in it.
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("first");
        builder.Writeln("second");
        builder.Writeln("third");
        builder.Writeln("fourth");

        var handler = new DeleteRangeWordTextHandler();
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            ["startParagraphIndex"] = 0,
            ["endParagraphIndex"] = 2,
            ["startCharIndex"] = 0,
            ["endCharIndex"] = 5
        });

        handler.Execute(CreateContext(document), parameters);

        Assert.DoesNotContain("second", document.GetText(), StringComparison.Ordinal);
    }
}
