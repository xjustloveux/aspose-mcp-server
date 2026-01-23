using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Link;
using AsposeMcpServer.Results.Pdf.Link;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Link;

public class GetPdfLinksHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfLinksHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Filter By Page

    [Fact]
    public void Execute_WithPageIndex_FiltersLinks()
    {
        var doc = CreateDocumentWithLinks(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLinksResult>(res);

        Assert.NotNull(result.PageIndex);
        Assert.Equal(1, result.PageIndex);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex", ex.Message);
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

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsLinks()
    {
        var doc = CreateDocumentWithLinks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLinksResult>(res);

        Assert.True(result.Count >= 0);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithLinks(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLinksResult>(res);

        Assert.Equal(3, result.Count);
    }

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var doc = CreateDocumentWithLinks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLinksResult>(res);

        Assert.Equal(2, result.Items.Count);
    }

    #endregion

    #region Link Details

    [Fact]
    public void Execute_ReturnsLinkIndex()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLinksResult>(res);
        var firstLink = result.Items[0];

        Assert.Equal(0, firstLink.Index);
    }

    [Fact]
    public void Execute_ReturnsPageIndex()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLinksResult>(res);
        var firstLink = result.Items[0];

        Assert.Equal(1, firstLink.PageIndex);
    }

    [Fact]
    public void Execute_ReturnsLinkType()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLinksResult>(res);
        var firstLink = result.Items[0];

        Assert.Equal("url", firstLink.Type);
    }

    #endregion

    #region No Links

    [Fact]
    public void Execute_NoLinks_ReturnsEmptyResult()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLinksResult>(res);

        Assert.Equal(0, result.Count);
        Assert.Contains("No links found", result.Message);
    }

    [Fact]
    public void Execute_NoLinksOnPage_ReturnsEmptyResult()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetLinksResult>(res);

        Assert.Equal(0, result.Count);
    }

    #endregion
}
