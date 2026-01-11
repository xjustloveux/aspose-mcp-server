using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Image;
using AsposeMcpServer.Tests.Helpers;
using DrawingColor = System.Drawing.Color;
using Rectangle = Aspose.Pdf.Rectangle;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Image;

[SupportedOSPlatform("windows")]
public class DeletePdfImageHandlerTests : PdfHandlerTestBase
{
    private readonly DeletePdfImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Basic Delete Operations

    [SkippableFact]
    public void Execute_DeletesImageFromDocument()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var doc = CreateDocumentWithImage();
        var imageCount = doc.Pages[1].Resources.Images.Count;
        Assert.True(imageCount > 0, "Document should have at least one image");

        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 1 },
            { "imageIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Deleted image", result);
        Assert.Contains("from page 1", result);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithDefaultParameters_DeletesFirstImageFromFirstPage()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
        var doc = CreateDocumentWithImage();
        var context = CreateContext(doc);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Deleted image 1", result);
        Assert.Contains("page 1", result);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithNoImages_ThrowsArgumentException()
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

    [Fact]
    public void Execute_WithInvalidPageIndex_ThrowsArgumentException()
    {
        var doc = CreateDocumentWithPages(2);
        var context = CreateContext(doc);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "pageIndex", 99 },
            { "imageIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("pageIndex must be between", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidImageIndex_ThrowsArgumentException()
    {
        SkipInEvaluationMode(AsposeLibraryType.Pdf);
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

    [Fact]
    public void Execute_WithPageIndexZero_UsesFirstPage()
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

    private Document CreateDocumentWithImage()
    {
        var doc = new Document();
        var page = doc.Pages.Add();
        var imagePath = CreateTestPngImage();
        page.AddImage(imagePath, new Rectangle(100, 600, 300, 800));
        return doc;
    }

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
}
