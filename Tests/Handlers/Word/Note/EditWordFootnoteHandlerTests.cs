using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class EditWordFootnoteHandlerTests : WordHandlerTestBase
{
    private readonly EditWordFootnoteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditFootnote()
    {
        Assert.Equal("edit_footnote", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsFootnote()
    {
        var doc = CreateDocumentWithFootnote();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Updated footnote text" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var footnotes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);
        Assert.Single(footnotes);
        if (!IsEvaluationMode()) Assert.Contains("Updated footnote text", footnotes[0].ToString(SaveFormat.Text));
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoteIndex_EditsSpecificFootnote()
    {
        var doc = CreateDocumentWithMultipleFootnotes();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "noteIndex", 1 },
            { "text", "Edited second footnote" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var footnotes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);
        Assert.Equal(2, footnotes.Count);
        if (!IsEvaluationMode())
        {
            Assert.Contains("First footnote", footnotes[0].ToString(SaveFormat.Text));
            Assert.Contains("Edited second footnote", footnotes[1].ToString(SaveFormat.Text));
        }

        AssertModified(context);
    }

    [Fact]
    public void Execute_WithReferenceMark_EditsFootnoteByMark()
    {
        var doc = CreateDocumentWithCustomMarkFootnote();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceMark", "*" },
            { "text", "Edited by reference mark" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var footnotes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);
        Assert.Single(footnotes);
        Assert.Equal("*", footnotes[0].ReferenceMark);
        if (!IsEvaluationMode()) Assert.Contains("Edited by reference mark", footnotes[0].ToString(SaveFormat.Text));
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFootnote();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoFootnotes_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("No footnotes here.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New footnote text" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFootnote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with footnote");
        builder.InsertFootnote(FootnoteType.Footnote, "Original footnote");
        return doc;
    }

    private static Document CreateDocumentWithMultipleFootnotes()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote");
        builder.Write(" more text");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote");
        return doc;
    }

    private static Document CreateDocumentWithCustomMarkFootnote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with custom footnote");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote with custom mark", "*");
        return doc;
    }

    #endregion
}
