using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Word.Text;
using AsposeMcpServer.Tests.Infrastructure;
using WordsShape = Aspose.Words.Drawing.Shape;
using WordsShapeType = Aspose.Words.Drawing.ShapeType;

namespace AsposeMcpServer.Tests.Handlers.Word.Text;

/// <summary>
///     Guards RB-11 and RB-20. Cross-paragraph deletion used to truncate only one run at each end,
///     so a paragraph built from several runs kept text the caller had asked to delete while the
///     operation still reported success. The same handler enumerated runs with a deep search, so an
///     inline text box inside the paragraph lost its content too.
/// </summary>
public class DeleteRangeMultiRunTests : TestBase
{
    /// <summary>
    ///     Builds a document whose paragraphs are each split across several runs.
    /// </summary>
    /// <param name="paragraphs">Run text per paragraph.</param>
    /// <returns>The document.</returns>
    private static Document BuildDocument(params string[][] paragraphs)
    {
        var doc = new Document();
        var body = doc.FirstSection.Body;
        body.RemoveAllChildren();

        foreach (var runs in paragraphs)
        {
            var para = new Aspose.Words.Paragraph(doc);
            foreach (var text in runs) para.AppendChild(new Run(doc, text));
            body.AppendChild(para);
        }

        return doc;
    }

    /// <summary>
    ///     Runs the delete_range operation over a document.
    /// </summary>
    /// <param name="doc">The document to modify.</param>
    /// <param name="startParagraph">Start paragraph index.</param>
    /// <param name="startChar">Start character index.</param>
    /// <param name="endParagraph">End paragraph index.</param>
    /// <param name="endChar">End character index.</param>
    private static void DeleteRange(Document doc, int startParagraph, int startChar,
        int endParagraph, int endChar)
    {
        var handler = new DeleteRangeWordTextHandler();
        var context = new OperationContext<Document> { Document = doc };
        var parameters = new OperationParameters();
        parameters.Set("startParagraphIndex", startParagraph);
        parameters.Set("startCharIndex", startChar);
        parameters.Set("endParagraphIndex", endParagraph);
        parameters.Set("endCharIndex", endChar);
        handler.Execute(context, parameters);
    }

    /// <summary>Reads a paragraph's own run text, excluding any inline shape content.</summary>
    /// <param name="doc">The document.</param>
    /// <param name="index">Paragraph index within the body.</param>
    /// <returns>The concatenated run text.</returns>
    private static string ParagraphText(Document doc, int index)
    {
        var para = doc.FirstSection.Body.Paragraphs[index];
        return string.Concat(para.GetChildNodes(NodeType.Run, true)
            .Cast<Run>()
            .Where(r => ReferenceEquals(r.ParentNode, para))
            .Select(r => r.Text));
    }

    [Fact]
    public void CrossParagraph_StartParagraphKeepsOnlyTextBeforeTheIndex()
    {
        var doc = BuildDocument(["AAA", "BBB", "CCC"], ["DDD", "EEE"]);

        DeleteRange(doc, 0, 4, 1, 4);

        Assert.Equal("AAAB", ParagraphText(doc, 0));
    }

    [Fact]
    public void CrossParagraph_EndParagraphKeepsOnlyTextAfterTheIndex()
    {
        var doc = BuildDocument(["AAA", "BBB", "CCC"], ["DDD", "EEE"]);

        DeleteRange(doc, 0, 4, 1, 4);

        Assert.Equal("EE", ParagraphText(doc, doc.FirstSection.Body.Paragraphs.Count - 1));
    }

    [Fact]
    public void CrossParagraph_EndIndexPastTheFirstRun_ShouldStillDelete()
    {
        var doc = BuildDocument(["AAA"], ["DDD", "EEE", "FFF"]);

        DeleteRange(doc, 0, 1, 1, 7);

        Assert.Equal("A", ParagraphText(doc, 0));
        Assert.Equal("FF", ParagraphText(doc, doc.FirstSection.Body.Paragraphs.Count - 1));
    }

    [Fact]
    public void CrossParagraph_MiddleParagraphsAreRemoved()
    {
        var doc = BuildDocument(["AAA"], ["MID"], ["ZZZ"]);
        var before = doc.FirstSection.Body.Paragraphs.Count;

        DeleteRange(doc, 0, 1, 2, 1);

        Assert.Equal(before - 1, doc.FirstSection.Body.Paragraphs.Count);
    }

    [Fact]
    public void InlineTextBoxContent_ShouldSurviveAnOuterParagraphDelete()
    {
        var doc = new Document();
        doc.FirstSection.Body.RemoveAllChildren();

        var para = new Aspose.Words.Paragraph(doc);
        para.AppendChild(new Run(doc, "OUTER-ONE"));
        var textBox = new WordsShape(doc, WordsShapeType.TextBox) { Width = 100, Height = 50 };
        var inner = new Aspose.Words.Paragraph(doc);
        inner.AppendChild(new Run(doc, "INSIDE"));
        textBox.AppendChild(inner);
        para.AppendChild(textBox);
        para.AppendChild(new Run(doc, "OUTER-TWO"));
        doc.FirstSection.Body.AppendChild(para);
        doc.FirstSection.Body.AppendChild(new Aspose.Words.Paragraph(doc));

        DeleteRange(doc, 0, 0, 0, 9);

        var shape = doc.GetChildNodes(NodeType.Shape, true).Cast<WordsShape>().Single();
        Assert.Contains("INSIDE", shape.GetText());
        Assert.Equal("OUTER-TWO", ParagraphText(doc, 0));
    }

    #region Range Ordering

    /// <summary>
    ///     A range whose end precedes its start is not a range. Only each index was bounds-checked
    ///     individually, so a reversed pair reached the deletion logic and either removed the wrong
    ///     span or silently removed nothing while still reporting success (R2-C04).
    /// </summary>
    /// <param name="startParagraph">Start paragraph index.</param>
    /// <param name="startChar">Start character index.</param>
    /// <param name="endParagraph">End paragraph index.</param>
    /// <param name="endChar">End character index.</param>
    [Theory]
    [InlineData(1, 0, 0, 0)]
    [InlineData(0, 3, 0, 1)]
    public void DeleteRange_WithAReversedRange_ShouldBeRejected(int startParagraph, int startChar,
        int endParagraph, int endChar)
    {
        var doc = BuildDocument(["Hello world"], ["Second paragraph"]);

        Assert.Throws<ArgumentException>(() =>
            DeleteRange(doc, startParagraph, startChar, endParagraph, endChar));
    }

    /// <summary>
    ///     A character index outside its paragraph cannot describe anything to delete.
    /// </summary>
    /// <param name="startChar">Start character index.</param>
    /// <param name="endChar">End character index.</param>
    [Theory]
    [InlineData(-1, 3)]
    [InlineData(0, -1)]
    [InlineData(0, 99)]
    [InlineData(99, 100)]
    public void DeleteRange_WithACharIndexOutsideTheParagraph_ShouldBeRejected(int startChar, int endChar)
    {
        var doc = BuildDocument(["Hello world"]);

        Assert.Throws<ArgumentException>(() => DeleteRange(doc, 0, startChar, 0, endChar));
    }

    [Fact]
    public void DeleteRange_WithAValidRange_ShouldStillDelete()
    {
        var doc = BuildDocument(["Hello world"]);

        DeleteRange(doc, 0, 0, 0, 6);

        Assert.Equal("world", ParagraphText(doc, 0));
    }

    /// <summary>
    ///     Validation counted <c>Paragraph.GetText()</c>, which includes the paragraph terminator,
    ///     while the deletion worked over the paragraph's own runs. An index one past the last real
    ///     character was therefore accepted and then deleted nothing (R3-C03).
    /// </summary>
    [Fact]
    public void DeleteRange_PastTheLastCharacter_ShouldBeRejected()
    {
        var doc = BuildDocument(["Hello"]);

        // "Hello" is five characters; GetText() reported six because of the terminator.
        var exception = Assert.Throws<ArgumentException>(() => DeleteRange(doc, 0, 5, 0, 6));

        Assert.Contains("5 character(s)", exception.Message);
        Assert.Equal("Hello", ParagraphText(doc, 0));
    }

    [Fact]
    public void DeleteRange_UpToTheLastCharacter_ShouldStillBeAccepted()
    {
        var doc = BuildDocument(["Hel", "lo"]);

        DeleteRange(doc, 0, 3, 0, 5);

        Assert.Equal("Hel", ParagraphText(doc, 0));
    }

    [Fact]
    public void DeleteRange_InAnEmptyParagraph_ShouldRejectAnyNonZeroIndex()
    {
        var doc = BuildDocument([], ["tail"]);

        var exception = Assert.Throws<ArgumentException>(() => DeleteRange(doc, 0, 0, 0, 1));

        Assert.Contains("0 character(s)", exception.Message);
    }

    [Fact]
    public void DeleteRange_AcrossParagraphs_ShouldRejectAnEndIndexPastTheLastCharacter()
    {
        var doc = BuildDocument(["first"], ["se", "cond"]);

        Assert.Throws<ArgumentException>(() => DeleteRange(doc, 0, 0, 1, 7));

        Assert.Equal("first", ParagraphText(doc, 0));
        Assert.Equal("second", ParagraphText(doc, 1));
    }

    #endregion
}
