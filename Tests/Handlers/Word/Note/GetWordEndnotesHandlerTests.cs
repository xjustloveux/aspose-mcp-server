using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Results.Word.Note;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class GetWordEndnotesHandlerTests : WordHandlerTestBase
{
    private readonly GetWordEndnotesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetEndnotes()
    {
        Assert.Equal("get_endnotes", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithEndnotes()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with endnote");
        builder.InsertFootnote(FootnoteType.Endnote, "First endnote");
        builder.Write(" and more text");
        builder.InsertFootnote(FootnoteType.Endnote, "Second endnote");
        return doc;
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_WithNoEndnotes_ReturnsEmptyList()
    {
        var doc = CreateDocumentWithText("Sample text without endnotes.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetWordNotesResult>(res);

        Assert.NotNull(result);
        Assert.Equal(0, result.Count);
        Assert.Contains("endnote", result.NoteType.ToLower());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithEndnotes_ReturnsEndnoteList()
    {
        var doc = CreateDocumentWithEndnotes();
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
