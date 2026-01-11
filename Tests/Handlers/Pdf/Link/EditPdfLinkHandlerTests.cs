using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Link;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Link;

public class EditPdfLinkHandlerTests : PdfHandlerTestBase
{
    private readonly EditPdfLinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Edit URL

    [Fact]
    public void Execute_WithUrl_ChangesLinkUrl()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 0 },
            { "url", "https://updated.com" }
        });

        _handler.Execute(context, parameters);

        var link = doc.Pages[1].Annotations.OfType<LinkAnnotation>().First();
        var action = link.Action as GoToURIAction;
        Assert.Equal("https://updated.com", action?.URI);
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithLinks(int count)
    {
        var doc = new Document();
        var page = doc.Pages.Add();

        for (var i = 0; i < count; i++)
        {
            var rect = new Rectangle(100, 700 - i * 30, 200, 720 - i * 30);
            var link = new LinkAnnotation(page, rect)
            {
                Action = new GoToURIAction($"https://example{i}.com")
            };
            page.Annotations.Add(link);
        }

        return doc;
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsLink()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 0 },
            { "url", "https://newurl.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Edited link", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsLinkIndex()
    {
        var doc = CreateDocumentWithLinks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 1 },
            { "url", "https://newurl.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("link 1", result);
    }

    [Fact]
    public void Execute_ReturnsPageIndex()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 0 },
            { "url", "https://newurl.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("page 1", result);
    }

    #endregion

    #region Edit Target Page

    [Fact]
    public void Execute_WithTargetPage_ChangesToInternalLink()
    {
        var doc = CreateDocumentWithLinks(1);
        doc.Pages.Add();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 0 },
            { "targetPage", 2 }
        });

        _handler.Execute(context, parameters);

        var link = doc.Pages[1].Annotations.OfType<LinkAnnotation>().First();
        Assert.IsType<GoToAction>(link.Action);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Execute_EditsAtVariousIndices(int linkIndex)
    {
        var doc = CreateDocumentWithLinks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", linkIndex },
            { "url", "https://newurl.com" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Edited link", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "linkIndex", 0 },
            { "url", "https://example.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutLinkIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "url", "https://example.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("linkIndex", ex.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Execute_WithPageIndexLessThanOne_ThrowsArgumentException(int invalidIndex)
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", invalidIndex },
            { "linkIndex", 0 },
            { "url", "https://example.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithPageIndexGreaterThanPageCount_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 99 },
            { "linkIndex", 0 },
            { "url", "https://example.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithLinkIndexOutOfRange_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 99 },
            { "url", "https://example.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("linkIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeLinkIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", -1 },
            { "url", "https://example.com" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("linkIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidTargetPage_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 0 },
            { "targetPage", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("targetPage", ex.Message);
    }

    #endregion
}
