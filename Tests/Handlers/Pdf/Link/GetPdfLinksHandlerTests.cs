using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Handlers.Pdf.Link;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("pageIndex", out var pageIndex));
        Assert.Equal(1, pageIndex.GetInt32());
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

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("count", out _));
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var doc = CreateDocumentWithLinks(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(3, json.RootElement.GetProperty("count").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsItemsArray()
    {
        var doc = CreateDocumentWithLinks(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(2, json.RootElement.GetProperty("items").GetArrayLength());
    }

    #endregion

    #region Link Details

    [Fact]
    public void Execute_ReturnsLinkIndex()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstLink = json.RootElement.GetProperty("items")[0];

        Assert.Equal(0, firstLink.GetProperty("index").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsPageIndex()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstLink = json.RootElement.GetProperty("items")[0];

        Assert.Equal(1, firstLink.GetProperty("pageIndex").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsLinkType()
    {
        var doc = CreateDocumentWithLinks(1);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstLink = json.RootElement.GetProperty("items")[0];

        Assert.Equal("url", firstLink.GetProperty("type").GetString());
    }

    #endregion

    #region No Links

    [Fact]
    public void Execute_NoLinks_ReturnsEmptyResult()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
        Assert.Contains("No links found", json.RootElement.GetProperty("message").GetString());
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

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion
}
