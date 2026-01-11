using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 0", result);
        Assert.Contains("endnote", result.ToLower());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithEndnotes_ReturnsEndnoteList()
    {
        var doc = CreateDocumentWithEndnotes();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("noteIndex", result);
        Assert.Contains("referenceMark", result);
        Assert.Contains("text", result);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var doc = CreateDocumentWithText("Sample text.");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("{", result);
        Assert.Contains("}", result);
        Assert.Contains("noteType", result);
    }

    #endregion
}
