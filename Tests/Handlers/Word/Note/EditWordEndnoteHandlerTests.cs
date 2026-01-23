using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Word.Note;

public class EditWordEndnoteHandlerTests : WordHandlerTestBase
{
    private readonly EditWordEndnoteHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_EditEndnote()
    {
        Assert.Equal("edit_endnote", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsEndnote()
    {
        var doc = CreateDocumentWithEndnote();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Updated endnote text" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Updated endnote text", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNoteIndex_EditsSpecificEndnote()
    {
        var doc = CreateDocumentWithMultipleEndnotes();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "noteIndex", 1 },
            { "text", "Edited second endnote" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("edited successfully", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutText_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithEndnote();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithNoEndnotes_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("No endnotes here.");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New endnote text" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithEndnote()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text with endnote");
        builder.InsertFootnote(FootnoteType.Endnote, "Original endnote");
        return doc;
    }

    private static Document CreateDocumentWithMultipleEndnotes()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Text");
        builder.InsertFootnote(FootnoteType.Endnote, "First endnote");
        builder.Write(" more text");
        builder.InsertFootnote(FootnoteType.Endnote, "Second endnote");
        return doc;
    }

    #endregion
}
