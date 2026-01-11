using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Content;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Content;

public class GetWordContentHandlerTests : WordHandlerTestBase
{
    private readonly GetWordContentHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetContent()
    {
        Assert.Equal("get_content", _handler.Operation);
    }

    #endregion

    #region Pagination Hints

    [Fact]
    public void Execute_WithMoreContent_ShowsContinueOffset()
    {
        var doc = CreateDocumentWithText("This is a longer text for testing pagination");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "maxChars", 10 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("use offset=", result);
    }

    #endregion

    #region Empty Document

    [Fact]
    public void Execute_WithEmptyDocument_ReturnsEmptyContent()
    {
        var doc = new Document();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Document Content", result);
    }

    #endregion

    #region Multiple Paragraphs

    [Fact]
    public void Execute_WithMultipleParagraphs_ReturnsAllContent()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph");
        builder.Writeln("Second paragraph");
        builder.Writeln("Third paragraph");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("First paragraph", result);
        Assert.Contains("Second paragraph", result);
        Assert.Contains("Third paragraph", result);
    }

    #endregion

    #region Basic Content Retrieval

    [Fact]
    public void Execute_ReturnsContent()
    {
        var doc = CreateDocumentWithText("Hello World");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Document Content", result);
    }

    [Fact]
    public void Execute_ReturnsDocumentText()
    {
        var doc = CreateDocumentWithText("Test content here");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Test content here", result);
    }

    [Fact]
    public void Execute_DoesNotMarkAsModified()
    {
        var doc = CreateDocumentWithText("Read only");
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.False(context.IsModified);
    }

    #endregion

    #region MaxChars Parameter

    [Fact]
    public void Execute_WithMaxChars_LimitsOutput()
    {
        var doc = CreateDocumentWithText("This is a longer text that should be truncated");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "maxChars", 10 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Showing chars", result);
    }

    [Fact]
    public void Execute_WithMaxChars_ShowsCharRange()
    {
        var doc = CreateDocumentWithText("Hello World Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "maxChars", 5 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Showing chars 0 to", result);
    }

    [Fact]
    public void Execute_WithMaxCharsLessThanTotal_ShowsMoreAvailable()
    {
        var doc = CreateDocumentWithText("This is a text that is longer than max chars");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "maxChars", 10 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("More content available", result);
    }

    #endregion

    #region Offset Parameter

    [Fact]
    public void Execute_WithOffset_SkipsCharacters()
    {
        var doc = CreateDocumentWithText("Hello World Test");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "offset", 6 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Document Content", result);
    }

    [Fact]
    public void Execute_WithOffsetAndMaxChars_CombinesParameters()
    {
        var doc = CreateDocumentWithText("Hello World Test Content");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "offset", 5 },
            { "maxChars", 10 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Showing chars 5 to", result);
    }

    [Fact]
    public void Execute_WithOffsetBeyondContent_ReturnsEmpty()
    {
        var doc = CreateDocumentWithText("Short");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "offset", 1000 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Document Content", result);
    }

    #endregion
}
