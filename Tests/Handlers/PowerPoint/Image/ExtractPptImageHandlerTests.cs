using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

// CA1416 - System.Drawing.Common is Windows-only, cross-platform support not required
#pragma warning disable CA1416
using System.Runtime.Versioning;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Image;

[SupportedOSPlatform("windows")]
[Collection("SerialSlides")]
public class ExtractPptImageHandlerTests : PptHandlerTestBase
{
    private readonly ExtractPptImageHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Extract()
    {
        SkipIfNotWindows();
        Assert.Equal("extract", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutPath_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Output Budget

    /// <summary>
    ///     The cap was applied to the total shape count, so a deck of text boxes holding no
    ///     pictures at all was refused for producing too many image files (R3-C05).
    /// </summary>
    [SkippableFact]
    public void Execute_WithManyTextShapesAndNoImages_ShouldNotBeRefused()
    {
        SkipIfNotWindows();
        var outputDir = Path.Combine(TestDir, "text_only_output");
        var pptxPath = Path.Combine(TestDir, "text_only.pptx");

        using (var authored = new Presentation())
        {
            var slide = authored.Slides[0];
            for (var i = 0; i <= RenderBudget.MaxOutputFiles; i++)
                slide.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 1, 1);

            authored.Save(pptxPath, SaveFormat.Pptx);
        }

        using var pres = new Presentation(pptxPath);
        var context = CreateContextWithPath(pres, pptxPath);
        var parameters = CreateEmptyParameters();
        parameters.Set("outputDir", outputDir);

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        Assert.Contains("Extracted 0 images", result.Message);
    }

    #endregion

    #region Basic Extract Operations

    [SkippableFact]
    public void Execute_ExtractsImagesFromPresentation()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithSkipDuplicates_SkipsDuplicateImages()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithJpegFormat_ExtractsAsJpeg()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithEmptyPresentation_ExtractsZeroImages()
    {
        SkipIfNotWindows();
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

    #region Duplicate Counting

    /// <summary>
    ///     The cap counted picture frames, but frames sharing an image produce one file between
    ///     them when duplicates are skipped, so a deck well inside the real output count was
    ///     refused (R4-R06).
    /// </summary>
    [SkippableFact]
    public void Execute_WithManyFramesSharingOneImage_ShouldNotBeRefusedWhenSkippingDuplicates()
    {
        SkipIfNotWindows();
        var outputDir = Path.Combine(TestDir, "duplicate_frames_output");
        var pptxPath = Path.Combine(TestDir, "duplicate_frames.pptx");

        using (var authored = new Presentation())
        {
            var slide = authored.Slides[0];
            using var bmp = new Bitmap(4, 4);
            using var ms = new MemoryStream();
            bmp.Save(ms, ImageFormat.Bmp);
            ms.Position = 0;
            var image = authored.Images.AddImage(ms);

            for (var i = 0; i <= RenderBudget.MaxOutputFiles; i++)
                slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 0, 0, 4, 4, image);

            authored.Save(pptxPath, SaveFormat.Pptx);
        }

        using var pres = new Presentation(pptxPath);
        var context = CreateContextWithPath(pres, pptxPath);
        var parameters = CreateEmptyParameters();
        parameters.Set("outputDir", outputDir);
        parameters.Set("skipDuplicates", true);

        var result = Assert.IsType<SuccessResult>(_handler.Execute(context, parameters));

        Assert.Contains("Extracted 1 images", result.Message);
        Assert.Single(Directory.GetFiles(outputDir));
    }

    /// <summary>
    ///     Without duplicate skipping every frame really does become a file, so the same deck is
    ///     still refused.
    /// </summary>
    [SkippableFact]
    public void Execute_WithManyFramesAndNoDuplicateSkipping_ShouldStillBeRefused()
    {
        SkipIfNotWindows();
        var outputDir = Path.Combine(TestDir, "duplicate_frames_kept_output");
        var pptxPath = Path.Combine(TestDir, "duplicate_frames_kept.pptx");

        using (var authored = new Presentation())
        {
            var slide = authored.Slides[0];
            using var bmp = new Bitmap(4, 4);
            using var ms = new MemoryStream();
            bmp.Save(ms, ImageFormat.Bmp);
            ms.Position = 0;
            var image = authored.Images.AddImage(ms);

            for (var i = 0; i <= RenderBudget.MaxOutputFiles; i++)
                slide.Shapes.AddPictureFrame(ShapeType.Rectangle, 0, 0, 4, 4, image);

            authored.Save(pptxPath, SaveFormat.Pptx);
        }

        using var pres = new Presentation(pptxPath);
        var context = CreateContextWithPath(pres, pptxPath);
        var parameters = CreateEmptyParameters();
        parameters.Set("outputDir", outputDir);

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
