using Aspose.Words;
using AsposeMcpServer.Handlers.Word.Hyperlink;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("This is a tooltip", hyperlinks[^1].ScreenTip);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://example.com", hyperlinks[^1].Address);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Equal("Click here", hyperlinks[^1].Result);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        if (!IsEvaluationMode(AsposeLibraryType.Words))
        {
            var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
            Assert.NotEmpty(hyperlinks);
            Assert.Equal("My Link", hyperlinks[^1].Result);
        }
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://example.com", hyperlinks[^1].Address);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://external.com", hyperlinks[^1].Address);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("MyBookmark", hyperlinks[^1].SubAddress);
    }

    #endregion

    #region Paragraph Index

    [Fact]
    public void Execute_WithParagraphIndex_InsertsAtPosition()
    {
        var doc = CreateDocumentWithParagraphs("First", "Second", "Third");
        var paragraphCountBefore = doc.GetChildNodes(NodeType.Paragraph, true).Count;
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "text", "New Link" },
            { "url", "https://example.com" },
            { "paragraphIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://example.com", hyperlinks[^1].Address);
        var paragraphCountAfter = doc.GetChildNodes(NodeType.Paragraph, true).Count;
        Assert.True(paragraphCountAfter > paragraphCountBefore);
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://example.com", hyperlinks[^1].Address);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) AssertContainsText(doc, "Start Link");
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);

        var hyperlinks = WordHyperlinkHelper.GetAllHyperlinks(doc);
        Assert.NotEmpty(hyperlinks);
        Assert.Equal("https://example.com", hyperlinks[^1].Address);
        if (!IsEvaluationMode(AsposeLibraryType.Words)) Assert.Equal("End Link", hyperlinks[^1].Result);
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
