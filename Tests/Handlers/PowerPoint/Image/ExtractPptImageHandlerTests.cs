using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

#pragma warning disable CA1416

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Image;

public class ExtractPptImageHandlerTests : PptHandlerTestBase
{
    private readonly ExtractPptImageHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Extract()
    {
        Assert.Equal("extract", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Extract Operations

    [Fact]
    public void Execute_ExtractsImagesFromPresentation()
    {
        var outputDir = Path.Combine(TestDir, "extract_output");
        var tempPptxPath = Path.Combine(TestDir, "test.pptx");

        var pres = CreatePresentationWithImage();
        pres.Save(tempPptxPath, SaveFormat.Pptx);
        pres.Dispose();

        pres = new Presentation(tempPptxPath);
        var context = CreateContextWithPath(pres, tempPptxPath);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("extracted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir);
        Assert.NotEmpty(files);
        foreach (var file in files)
        {
            var fileInfo = new FileInfo(file);
            Assert.True(fileInfo.Length > 0, $"Extracted file {file} should have content");
        }
    }

    [Fact]
    public void Execute_WithSkipDuplicates_SkipsDuplicateImages()
    {
        var outputDir = Path.Combine(TestDir, "extract_skip");
        var tempPptxPath = Path.Combine(TestDir, "test_dup.pptx");

        var pres = CreatePresentationWithDuplicateImages();
        pres.Save(tempPptxPath, SaveFormat.Pptx);
        pres.Dispose();

        pres = new Presentation(tempPptxPath);
        var context = CreateContextWithPath(pres, tempPptxPath);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "skipDuplicates", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("extracted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithJpegFormat_ExtractsAsJpeg()
    {
        var outputDir = Path.Combine(TestDir, "extract_jpeg");
        var tempPptxPath = Path.Combine(TestDir, "test_jpeg.pptx");

        var pres = CreatePresentationWithImage();
        pres.Save(tempPptxPath, SaveFormat.Pptx);
        pres.Dispose();

        pres = new Presentation(tempPptxPath);
        var context = CreateContextWithPath(pres, tempPptxPath);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "format", "jpeg" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("extracted", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithEmptyPresentation_ExtractsZeroImages()
    {
        var outputDir = Path.Combine(TestDir, "extract_empty");
        var tempPptxPath = Path.Combine(TestDir, "test_empty.pptx");

        var pres = CreateEmptyPresentation();
        pres.Save(tempPptxPath, SaveFormat.Pptx);
        pres.Dispose();

        pres = new Presentation(tempPptxPath);
        var context = CreateContextWithPath(pres, tempPptxPath);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("extracted 0 images", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithImage()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];

        using var bmp = new Bitmap(100, 100);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Cyan);

        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Bmp);
        ms.Position = 0;

        var image = pres.Images.AddImage(ms);
        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image);

        return pres;
    }

    private static Presentation CreatePresentationWithDuplicateImages()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];

        using var bmp = new Bitmap(100, 100);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Magenta);

        using var ms = new MemoryStream();
        bmp.Save(ms, ImageFormat.Bmp);
        ms.Position = 0;

        var image = pres.Images.AddImage(ms);
        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image);
        slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 200, 50, 100, 100, image);

        return pres;
    }

    #endregion
}
