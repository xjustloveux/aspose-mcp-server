using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Image;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Image;

public class EditPdfImageHandlerTests : PdfHandlerTestBase
{
    private readonly EditPdfImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Error Handling - No Images

    [Fact]
    public void Execute_WithNoImagesOnPage_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "imageIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("imageIndex must be between", ex.Message);
    }

    #endregion

    #region Error Handling - Image Path

    [Fact]
    public void Execute_WithNonExistentImagePath_ThrowsFileNotFoundException()
    {
        var doc = CreateDocumentWithImage();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "imageIndex", 1 },
            { "imagePath", "nonexistent_file_12345.png" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Error Handling - Page Index

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 99 },
            { "imageIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Execute_WithPageIndexZero_UsesPageOne()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 0 },
            { "imageIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("imageIndex must be between", ex.Message);
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
