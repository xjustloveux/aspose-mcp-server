using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Image;

public class ExportSlidesHandlerTests : PptHandlerTestBase
{
    private readonly ExportSlidesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ExportSlides()
    {
        Assert.Equal("export_slides", _handler.Operation);
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

    #region Basic Export Operations

    [Fact]
    public void Execute_ExportsSlidesToImages()
    {
        var outputDir = Path.Combine(TestDir, "export_output");
        var tempPptxPath = Path.Combine(TestDir, "test.pptx");

        var pres = CreatePresentationWithSlides(2);
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

        Assert.Contains("exported", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir);
        Assert.Equal(2, files.Length);
        foreach (var file in files)
        {
            var fileInfo = new FileInfo(file);
            Assert.True(fileInfo.Length > 0, $"Exported file {file} should have content");
        }
    }

    [Fact]
    public void Execute_WithSpecificSlideIndexes_ExportsSelectedSlides()
    {
        var outputDir = Path.Combine(TestDir, "export_selected");
        var tempPptxPath = Path.Combine(TestDir, "test_selected.pptx");

        var pres = CreatePresentationWithSlides(3);
        pres.Save(tempPptxPath, SaveFormat.Pptx);
        pres.Dispose();

        pres = new Presentation(tempPptxPath);
        var context = CreateContextWithPath(pres, tempPptxPath);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outputDir },
            { "slideIndexes", "0,1" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("exported", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir);
        Assert.Equal(2, files.Length);
    }

    [Fact]
    public void Execute_WithJpegFormat_ExportsAsJpeg()
    {
        var outputDir = Path.Combine(TestDir, "export_jpeg");
        var tempPptxPath = Path.Combine(TestDir, "test_jpeg.pptx");

        var pres = CreateEmptyPresentation();
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

        Assert.Contains("exported", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir, "*.jpg");
        Assert.NotEmpty(files);
    }

    #endregion
}
