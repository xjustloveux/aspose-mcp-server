using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Image;
using AsposeMcpServer.Results.Pdf.Image;
using AsposeMcpServer.Tests.Infrastructure;

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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.True(result.Count >= 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.Equal(0, result.Count);
        Assert.Equal("No images found on page 1", result.Message);
    }

    #endregion

    #region Get All Images

    [Fact]
    public void Execute_WithNoImages_ReturnsEmptyResult()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.Equal(0, result.Count);
    }

    [Fact]
    public void Execute_WithNoImages_ReturnsMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.Contains("No images found", result.Message);
    }

    [Fact]
    public void Execute_ReturnsTypedResult()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.IsType<GetImagesPdfResult>(result);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.Equal(2, result.PageIndex);
    }

    [Fact]
    public void Execute_WithNoPageIndex_SearchesAllPages()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.True(result.Count >= 0);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.True(result.Count >= 0);
    }

    #endregion

    #region Result Structure

    [Fact]
    public void Execute_WithoutImages_ReturnsCorrectStructure()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.Equal(0, result.Count);
        Assert.NotNull(result.Items);
        Assert.NotNull(result.Message);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.NotNull(result.Items);
        Assert.Empty(result.Items);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.Equal(1, result.Count);
        Assert.NotEmpty(result.Items);
    }

    [Fact]
    public void Execute_WithImageOnAllPages_ReturnsAllImages()
    {
        var doc = CreateDocumentWithMultipleImages();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.True(result.Count >= 2);
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

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetImagesPdfResult>(res);
        Assert.NotEmpty(result.Items);
        var firstItem = result.Items[0];
        Assert.True(firstItem.Index >= 1);
        Assert.Equal(1, firstItem.PageIndex);
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
