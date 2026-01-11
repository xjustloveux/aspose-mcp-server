using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class DeleteWordEndnoteHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordEndnoteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteEndnote()
    {
        Assert.Equal("delete_endnote", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidNoteIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithEndnotes();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "noteIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
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

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesAllEndnotes()
    {
        var doc = CreateDocumentWithEndnotes();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        Assert.Contains("2", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoteIndex_DeletesSpecificEndnote()
    {
        var doc = CreateDocumentWithEndnotes();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "noteIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        Assert.Contains("1", result);
    }

    #endregion
}
