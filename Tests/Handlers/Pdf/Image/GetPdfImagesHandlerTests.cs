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
}
