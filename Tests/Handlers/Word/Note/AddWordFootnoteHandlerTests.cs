using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class AddWordFootnoteHandlerTests : WordHandlerTestBase
{
    private readonly AddWordFootnoteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddFootnote()
    {
        Assert.Equal("add_footnote", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsFootnote()
    {
        var doc = CreateDocumentWithText("Sample text for footnote.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "This is a footnote" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var footnotes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);
        Assert.Single(footnotes);
        if (!IsEvaluationMode()) Assert.Contains("This is a footnote", footnotes[0].ToString(SaveFormat.Text));
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithParagraphIndex_AddsFootnoteAtParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph", "Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Footnote at paragraph" },
            { "paragraphIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var footnotes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);
        Assert.Single(footnotes);
        if (!IsEvaluationMode()) Assert.Contains("Footnote at paragraph", footnotes[0].ToString(SaveFormat.Text));
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomMark_AddsFootnoteWithCustomMark()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Custom mark footnote" },
            { "customMark", "*" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var footnotes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Footnote);
        Assert.Single(footnotes);
        Assert.Equal("*", footnotes[0].ReferenceMark);
        if (!IsEvaluationMode()) Assert.Contains("Custom mark footnote", footnotes[0].ToString(SaveFormat.Text));
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSectionIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Footnote" },
            { "sectionIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
