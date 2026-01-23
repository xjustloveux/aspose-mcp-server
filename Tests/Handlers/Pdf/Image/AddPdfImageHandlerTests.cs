using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.Pdf.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using DrawingColor = System.Drawing.Color;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Image;

[SupportedOSPlatform("windows")]
public class AddPdfImageHandlerTests : PdfHandlerTestBase
{
    private readonly AddPdfImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Helper Methods

    private string CreateTestPngImage()
    {
        var imagePath = Path.Combine(TestDir, $"test_{Guid.NewGuid()}.png");
        using var bitmap = new Bitmap(100, 100);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(DrawingColor.Red);
        }

        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsImageToDocument()
    {
        var imagePath = CreateTestPngImage();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added image", result.Message);
        Assert.Contains("page 1", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_AddsImageToSpecificPage()
    {
        var imagePath = CreateTestPngImage();
        var doc = CreateDocumentWithPages(3);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "pageIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added image", result.Message);
        Assert.Contains("page 2", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomPosition_AddsImage()
    {
        var imagePath = CreateTestPngImage();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "x", 200.0 },
            { "y", 400.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added image", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithCustomSize_AddsImage()
    {
        var imagePath = CreateTestPngImage();
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "width", 150.0 },
            { "height", 100.0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Added image", result.Message);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutImagePath_ThrowsArgumentException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("imagePath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        var doc = CreateEmptyDocument();
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", "/nonexistent/path/image.png" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var imagePath = CreateTestPngImage();
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "pageIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [Fact]
    public void Execute_WithPageIndexZero_AddsToFirstPage()
    {
        var imagePath = CreateTestPngImage();
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "pageIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("page 1", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithNegativePageIndex_AddsToFirstPage()
    {
        var imagePath = CreateTestPngImage();
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "imagePath", imagePath },
            { "pageIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("page 1", result.Message);
        AssertModified(context);
    }

    #endregion
}
