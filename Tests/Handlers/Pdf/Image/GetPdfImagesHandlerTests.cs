using System.Text.Json.Nodes;
using AsposeMcpServer.Handlers.Pdf.Image;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Image;

public class GetPdfImagesHandlerTests : PdfHandlerTestBase
{
    private readonly GetPdfImagesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Negative Page Index

    [Fact]
    public void Execute_WithNegativePageIndex_SearchesAllPages()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", -1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result);
        Assert.DoesNotContain("pageIndex must be between", result);
    }

    #endregion

    #region Specific Page With No Images

    [Fact]
    public void Execute_WithSpecificPageNoImages_ReturnsEmptyForPage()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("\"count\": 0", result);
        Assert.Contains("No images found on page 1", result);
    }

    #endregion

    #region Get All Images

    [Fact]
    public void Execute_WithNoImages_ReturnsEmptyResult()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result);
        Assert.Contains("\"count\": 0", result);
    }

    [Fact]
    public void Execute_WithNoImages_ReturnsMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("No images found", result);
    }

    [Fact]
    public void Execute_ReturnsJsonFormat()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
    }

    #endregion

    #region Page Index Parameter

    [Fact]
    public void Execute_WithPageIndex_ReturnsPageInfo()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("pageIndex", result);
        Assert.Contains("2", result);
    }

    [Fact]
    public void Execute_WithNoPageIndex_SearchesAllPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Execute_WithPageIndexZero_SearchesAllPages()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("count", result);
        Assert.DoesNotContain("pageIndex must be between", result);
    }

    #endregion

    #region Result Structure

    [Fact]
    public void Execute_WithoutImages_ReturnsCorrectStructure()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.NotNull(json["count"]);
        Assert.NotNull(json["items"]);
        Assert.NotNull(json["message"]);
    }

    [Fact]
    public void Execute_WithPageIndex_ReturnsItemsArray()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.NotNull(json["items"]);
        Assert.Empty(json["items"]!.AsArray());
    }

    #endregion
}
