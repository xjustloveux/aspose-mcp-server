using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class DeleteWordFootnoteHandlerTests : WordHandlerTestBase
{
    private readonly DeleteWordFootnoteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_DeleteFootnote()
    {
        Assert.Equal("delete_footnote", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidNoteIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithFootnotes();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "noteIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
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

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesAllFootnotes()
    {
        var doc = CreateDocumentWithFootnotes();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("deleted", result.ToLower());
        Assert.Contains("2", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoteIndex_DeletesSpecificFootnote()
    {
        var doc = CreateDocumentWithFootnotes();
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
