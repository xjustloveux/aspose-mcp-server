using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

    private static Document CreateDocumentWithCustomMarkFootnote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with custom footnote");
        builder.InsertFootnote(FootnoteType.Footnote, "Footnote with custom mark", "*");
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("1", result.Message);
    }

    [Fact]
    public void Execute_WithReferenceMark_DeletesFootnoteByMark()
    {
        var doc = CreateDocumentWithCustomMarkFootnote();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "referenceMark", "*" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("1", result.Message);
        AssertModified(context);
    }

    #endregion
}
