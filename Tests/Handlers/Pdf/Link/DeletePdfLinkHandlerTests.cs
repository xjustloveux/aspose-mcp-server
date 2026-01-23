using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Link;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Link;

public class DeletePdfLinkHandlerTests : PdfHandlerTestBase
{
    private readonly DeletePdfLinkHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
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

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesLink()
    {
        var doc = CreateDocumentWithLinks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
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
            { "linkIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Link 1", result.Message);
    }

    [Fact]
    public void Execute_ReturnsPageIndex()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("page 1", result.Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_DeletesAtVariousIndices(int linkIndex)
    {
        var doc = CreateDocumentWithLinks(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", linkIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
    }

    [Fact]
    public void Execute_ReducesLinkCount()
    {
        var doc = CreateDocumentWithLinks(3);
        var initialCount = doc.Pages[1].Annotations.OfType<LinkAnnotation>().Count();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 0 }
        });

        _handler.Execute(context, parameters);

        var finalCount = doc.Pages[1].Annotations.OfType<LinkAnnotation>().Count();
        Assert.Equal(initialCount - 1, finalCount);
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
            { "linkIndex", 0 }
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
            { "pageIndex", 1 }
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
            { "linkIndex", 0 }
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
            { "linkIndex", 0 }
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
            { "linkIndex", 99 }
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
            { "linkIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("linkIndex", ex.Message);
    }

    [Fact]
    public void Execute_NoLinks_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "linkIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("linkIndex", ex.Message);
    }

    #endregion
}
