using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Hyperlink;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://new.com", hyperlinks[0].Address);
        Assert.Equal("New Tooltip", hyperlinks[0].ScreenTip);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Equal("New Text", hyperlinks[0].Result);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://newurl.com", hyperlinks[0].Address);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
            Assert.NotEmpty(hyperlinks);
            Assert.Equal("New Text", hyperlinks[0].Result);
        }
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://updated.com", hyperlinks[0].Address);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
            Assert.NotEmpty(hyperlinks);
            Assert.Equal("Updated Link Text", hyperlinks[0].Result);
        }
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("New tooltip", hyperlinks[0].ScreenTip);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("Bookmark1", hyperlinks[0].SubAddress);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://original.com", hyperlinks[0].Address);
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
