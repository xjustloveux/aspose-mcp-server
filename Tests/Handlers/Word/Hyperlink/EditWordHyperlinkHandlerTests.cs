using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Hyperlink;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Word.Hyperlink;

public class EditWordHyperlinkHandlerTests : WordHandlerTestBase
{
    private readonly EditWordHyperlinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Multiple Changes

    [Fact]
    public void Execute_WithMultipleChanges_AppliesAll()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "url", "https://new.com" },
            { "displayText", "New Text" },
            { "tooltip", "New Tooltip" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("URL: https://new.com", result);
        Assert.Contains("Display text: New Text", result);
        Assert.Contains("Tooltip: New Tooltip", result);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithHyperlink()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertHyperlink("Original Link", "https://original.com", false);
        return doc;
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsHyperlink()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "url", "https://newurl.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("edited successfully", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsHyperlinkIndex()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "displayText", "New Text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("#0", result);
    }

    #endregion

    #region Edit Properties

    [Fact]
    public void Execute_WithUrl_ChangesUrl()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "url", "https://updated.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("URL: https://updated.com", result);
    }

    [Fact]
    public void Execute_WithDisplayText_ChangesDisplayText()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "displayText", "Updated Link Text" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Display text: Updated Link Text", result);
    }

    [Fact]
    public void Execute_WithTooltip_ChangesTooltip()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "tooltip", "New tooltip" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Tooltip: New tooltip", result);
    }

    [Fact]
    public void Execute_WithSubAddress_ChangesSubAddress()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "subAddress", "Bookmark1" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("SubAddress: Bookmark1", result);
    }

    [Fact]
    public void Execute_WithNoChanges_ReturnsNoChangeMessage()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("No change parameters provided", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidHyperlinkIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 99 },
            { "url", "https://example.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeHyperlinkIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithHyperlink();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", -1 },
            { "url", "https://example.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void Execute_NoHyperlinks_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithText("No hyperlinks");
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hyperlinkIndex", 0 },
            { "url", "https://example.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("no hyperlinks", ex.Message);
    }

    #endregion
}
