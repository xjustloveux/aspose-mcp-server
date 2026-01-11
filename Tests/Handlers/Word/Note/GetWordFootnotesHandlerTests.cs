using Aspose.Words;
using Aspose.Words.Notes;
using AsposeMcpServer.Handlers.Word.Note;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 0", result);
        Assert.Contains("footnote", result.ToLower());
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithFootnotes_ReturnsFootnoteList()
    {
        var doc = CreateDocumentWithFootnotes();
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
