using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Image;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Image;

public class ExtractPdfImageHandlerTests : PdfHandlerTestBase
{
    private readonly ExtractPdfImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Extract()
    {
        Assert.Equal("extract", _handler.Operation);
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

    #endregion

    #region Image Index Error Handling

    [Fact]
    public void Execute_WithInvalidImageIndex_OnPageWithImages_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithImage();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "imageIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("imageIndex must be between", ex.Message);
    }

    #endregion

    #region Basic Extract Operations

    [Fact]
    public void Execute_WithNoImages_ReturnsNoImagesMessage()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("No images found", result);
    }

    [Fact]
    public void Execute_DefaultPageIndex_UsesPageOne()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("page 1", result);
    }

    #endregion

    #region Page Index Parameter

    [Fact]
    public void Execute_WithPageIndex_ExtractsFromCorrectPage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 2 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("page 2", result);
    }

    [Fact]
    public void Execute_WithPageIndexZero_ReturnsNoImagesMessage()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("No images found", result);
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
