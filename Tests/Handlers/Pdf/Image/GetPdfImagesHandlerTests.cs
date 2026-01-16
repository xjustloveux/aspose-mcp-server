using System.Text.Json.Nodes;
using Aspose.Pdf;
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

    #region PDF With Images

    [Fact]
    public void Execute_WithImageOnSpecificPage_ReturnsImageInfo()
    {
        var doc = CreateDocumentWithImage();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.Equal(1, json["count"]!.GetValue<int>());
        Assert.NotEmpty(json["items"]!.AsArray());
    }

    [Fact]
    public void Execute_WithImageOnAllPages_ReturnsAllImages()
    {
        var doc = CreateDocumentWithMultipleImages();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        Assert.True(json["count"]!.GetValue<int>() >= 2);
    }

    [Fact]
    public void Execute_WithImageInfo_ReturnsWidthHeight()
    {
        var doc = CreateDocumentWithImage();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonNode.Parse(result);
        Assert.NotNull(json);
        var items = json["items"]!.AsArray();
        Assert.NotEmpty(items);
        var firstItem = items[0];
        Assert.NotNull(firstItem);
        Assert.True(firstItem["index"]!.GetValue<int>() >= 1);
        Assert.Equal(1, firstItem["pageIndex"]!.GetValue<int>());
    }

    #endregion

    #region Helper Methods

    private static Document CreateDocumentWithImage()
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        var rect = new Rectangle(100, 600, 300, 800);
        using var imageStream = CreateSimpleImageStream();
        page.AddImage(imageStream, rect);
        return doc;
    }

    private static Document CreateDocumentWithMultipleImages()
    {
        var doc = new Document();
        var page1 = doc.Pages.Add();
        var page2 = doc.Pages.Add();
        var rect1 = new Rectangle(100, 600, 300, 800);
        var rect2 = new Rectangle(100, 400, 300, 600);
        using (var imageStream1 = CreateSimpleImageStream())
        {
            page1.AddImage(imageStream1, rect1);
        }

        using (var imageStream2 = CreateSimpleImageStream())
        {
            page2.AddImage(imageStream2, rect2);
        }

        return doc;
    }

    private static MemoryStream CreateSimpleImageStream()
    {
        var ms = new MemoryStream();
        var width = 100;
        var height = 100;
        var bmp = new byte[width * height * 3 + 54];
        bmp[0] = 0x42;
        bmp[1] = 0x4D;
        var fileSize = bmp.Length;
        bmp[2] = (byte)(fileSize & 0xFF);
        bmp[3] = (byte)((fileSize >> 8) & 0xFF);
        bmp[4] = (byte)((fileSize >> 16) & 0xFF);
        bmp[5] = (byte)((fileSize >> 24) & 0xFF);
        bmp[10] = 54;
        bmp[14] = 40;
        bmp[18] = (byte)(width & 0xFF);
        bmp[19] = (byte)((width >> 8) & 0xFF);
        bmp[22] = (byte)(height & 0xFF);
        bmp[23] = (byte)((height >> 8) & 0xFF);
        bmp[26] = 1;
        bmp[28] = 24;
        for (var i = 54; i < bmp.Length; i += 3)
        {
            bmp[i] = 255;
            bmp[i + 1] = 255;
            bmp[i + 2] = 255;
        }

        ms.Write(bmp, 0, bmp.Length);
        ms.Position = 0;
        return ms;
    }

    #endregion
}
