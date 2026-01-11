using AsposeMcpServer.Handlers.Word.Hyperlink;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Hyperlink;

public class AddWordHyperlinkHandlerTests : WordHandlerTestBase
{
    private readonly AddWordHyperlinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Tooltip

    [Fact]
    public void Execute_WithTooltip_SetsTooltip()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Link" },
            { "url", "https://example.com" },
            { "tooltip", "This is a tooltip" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Tooltip:", result);
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsHyperlink()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Click here" },
            { "url", "https://example.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added successfully", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsDisplayText()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "My Link" },
            { "url", "https://example.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("My Link", result);
    }

    [Fact]
    public void Execute_ReturnsUrl()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Link" },
            { "url", "https://example.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("https://example.com", result);
    }

    #endregion

    #region URL vs SubAddress

    [Fact]
    public void Execute_WithUrl_AddsUrlHyperlink()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "External Link" },
            { "url", "https://external.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("URL:", result);
    }

    [Fact]
    public void Execute_WithSubAddress_AddsBookmarkHyperlink()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Bookmark Link" },
            { "subAddress", "MyBookmark" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("SubAddress", result);
    }

    #endregion

    #region Paragraph Index

    [Fact]
    public void Execute_WithParagraphIndex_InsertsAtPosition()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New Link" },
            { "url", "https://example.com" },
            { "paragraphIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("after paragraph #1", result);
    }

    [Fact]
    public void Execute_WithParagraphIndexMinusOne_InsertsAtBeginning()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Start Link" },
            { "url", "https://example.com" },
            { "paragraphIndex", -1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("beginning of document", result);
    }

    [Fact]
    public void Execute_WithoutParagraphIndex_InsertsAtEnd()
    {
        var doc = CreateDocumentWithText("Some text");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "End Link" },
            { "url", "https://example.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("end of document", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutUrlOrSubAddress_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Link without target" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("url", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidParagraphIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithParagraphs("Only one");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "Link" },
            { "url", "https://example.com" },
            { "paragraphIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion
}
