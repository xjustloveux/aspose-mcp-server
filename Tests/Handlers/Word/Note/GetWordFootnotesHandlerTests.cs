using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Results.Word.Note;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class GetWordFootnotesHandlerTests : WordHandlerTestBase
{
    private readonly GetWordFootnotesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetFootnotes()
    {
        Assert.Equal("get_footnotes", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithFootnotes()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with footnote");
        builder.InsertFootnote(FootnoteType.Footnote, "First footnote");
        builder.Write(" and more text");
        builder.InsertFootnote(FootnoteType.Footnote, "Second footnote");
        return doc;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoFootnotes_ReturnsEmptyList()
    {
        var doc = CreateDocumentWithText("Sample text without footnotes.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordNotesResult>(res);

        Assert.NotNull(result);
        Assert.Equal(0, result.Count);
        Assert.Contains("footnote", result.NoteType.ToLower());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithFootnotes_ReturnsFootnoteList()
    {
        var doc = CreateDocumentWithFootnotes();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordNotesResult>(res);

        Assert.NotNull(result);
        Assert.True(result.Count >= 2);
        Assert.NotNull(result.Notes);
        Assert.NotEmpty(result.Notes);
        Assert.True(result.Notes[0].NoteIndex >= 0);
    }

    [Fact]
    public void Execute_ReturnsNoteType()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordNotesResult>(res);

        Assert.NotNull(result);
        Assert.NotNull(result.NoteType);
    }

    #endregion
}
