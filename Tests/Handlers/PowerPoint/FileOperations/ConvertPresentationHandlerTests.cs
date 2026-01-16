using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.FileOperations;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.FileOperations;

public class ConvertPresentationHandlerTests : PptHandlerTestBase
{
    private readonly ConvertPresentationHandler _handler = new();
    private readonly string _inputPath;

    public ConvertPresentationHandlerTests()
    {
        _inputPath = Path.Combine(TestDir, "input.pptx");

        using var pres = new Presentation();
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 100);
        pres.Save(_inputPath, SaveFormat.Pptx);
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_Convert()
    {
        Assert.Equal("convert", _handler.Operation);
    }

    #endregion

    #region Basic Convert Operations

    [Fact]
    public void Execute_ConvertsToPdf()
    {
        var outputPath = Path.Combine(TestDir, "output.pdf");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputPath", outputPath },
            { "format", "pdf" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("converted", result.ToLower());
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted PDF file should have content");
    }

    [Fact]
    public void Execute_ConvertsToHtml()
    {
        var outputPath = Path.Combine(TestDir, "output.html");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputPath", outputPath },
            { "format", "html" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("converted", result.ToLower());
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted HTML file should have content");
    }

    [Fact]
    public void Execute_ConvertsToPpt()
    {
        var outputPath = Path.Combine(TestDir, "output.ppt");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputPath", outputPath },
            { "format", "ppt" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("converted", result.ToLower());
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Converted PPT file should have content");

        using var convertedPres = new Presentation(outputPath);
        Assert.True(convertedPres.Slides.Count > 0, "Converted presentation should have slides");
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSource_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "output.pdf") },
            { "format", "pdf" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputPath_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "format", "pdf" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutFormat_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputPath", Path.Combine(TestDir, "output.pdf") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedFormat_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputPath", Path.Combine(TestDir, "output.xyz") },
            { "format", "xyz" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Additional Format Tests

    [Theory]
    [InlineData("pptx")]
    [InlineData("odp")]
    [InlineData("xps")]
    [InlineData("tiff")]
    public void Execute_WithVariousFormats_Converts(string format)
    {
        var outputPath = Path.Combine(TestDir, $"output.{format}");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputPath", outputPath },
            { "format", format }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("converted", result.ToLower());
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_WithJpegFormat_ConvertsSlide()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode limits image conversion");

        var outputPath = Path.Combine(TestDir, "slide.jpg");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputPath", outputPath },
            { "format", "jpeg" },
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Slide 0", result);
        Assert.Contains("JPEG", result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_WithPngFormat_ConvertsSlide()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode limits image conversion");

        var outputPath = Path.Combine(TestDir, "slide.png");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputPath", outputPath },
            { "format", "png" },
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Slide 0", result);
        Assert.Contains("PNG", result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Execute_WithJpgFormat_ConvertsSlide()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode limits image conversion");

        var outputPath = Path.Combine(TestDir, "slide_jpg.jpg");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputPath", outputPath },
            { "format", "jpg" },
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("JPEG", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion
}
