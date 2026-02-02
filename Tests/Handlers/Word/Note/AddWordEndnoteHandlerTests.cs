using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class AddWordEndnoteHandlerTests : WordHandlerTestBase
{
    private readonly AddWordEndnoteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_AddEndnote()
    {
        Assert.Equal("add_endnote", _handler.Operation);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsEndnote()
    {
        var doc = CreateDocumentWithText("Sample text for endnote.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "This is an endnote" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var endnotes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Endnote);
        Assert.Single(endnotes);
        if (!IsEvaluationMode()) Assert.Contains("This is an endnote", endnotes[0].ToString(SaveFormat.Text));
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithParagraphIndex_AddsEndnoteAtParagraph()
    {
        var doc = CreateDocumentWithParagraphs("First paragraph", "Second paragraph", "Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Endnote at paragraph" },
            { "paragraphIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var endnotes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Endnote);
        Assert.Single(endnotes);
        if (!IsEvaluationMode()) Assert.Contains("Endnote at paragraph", endnotes[0].ToString(SaveFormat.Text));
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomMark_AddsEndnoteWithCustomMark()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Custom mark endnote" },
            { "customMark", "†" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var endnotes = WordNoteHelper.GetNotesFromDoc(doc, FootnoteType.Endnote);
        Assert.Single(endnotes);
        Assert.Equal("†", endnotes[0].ReferenceMark);
        if (!IsEvaluationMode()) Assert.Contains("Custom mark endnote", endnotes[0].ToString(SaveFormat.Text));
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
            { "text", "Endnote" },
            { "sectionIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
