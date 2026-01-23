using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.FileOperations;

public class SplitPresentationHandlerTests : PptHandlerTestBase
{
    private readonly SplitPresentationHandler _handler = new();
    private readonly string _inputPath;

    public SplitPresentationHandlerTests()
    {
        _inputPath = Path.Combine(TestDir, "input.pptx");

        using var pres = new Presentation();
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 100);
        pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
        pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
        pres.Save(_inputPath, SaveFormat.Pptx);
    }

    #region Operation Property

    [Fact]
    public void Operation_Returns_Split()
    {
        Assert.Equal("split", _handler.Operation);
    }

    #endregion

    #region Basic Split Operations

    [SkippableFact]
    public void Execute_SplitsPresentation()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_output");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("3", result.Message);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir, "*.pptx");
        Assert.Equal(3, files.Length);
        foreach (var file in files)
        {
            var fileInfo = new FileInfo(file);
            Assert.True(fileInfo.Length > 0, $"Split file {file} should have content");
        }
    }

    [SkippableFact]
    public void Execute_WithPath_SplitsPresentation()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_path");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", _inputPath },
            { "outputDirectory", outputDir }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(Directory.Exists(outputDir));
        var files = Directory.GetFiles(outputDir, "*.pptx");
        Assert.NotEmpty(files);
    }

    [SkippableFact]
    public void Execute_WithSlidesPerFile_SplitsPresentation()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_multi");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "slidesPerFile", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("2", result.Message);
    }

    [SkippableFact]
    public void Execute_WithOutputFileNamePattern_UsesPattern()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Master slide cloning fails in evaluation mode");

        var outputDir = Path.Combine(TestDir, "split_pattern");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "outputFileNamePattern", "presentation_{index}.pptx" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("split", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(Path.Combine(outputDir, "presentation_0.pptx")));
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
            { "outputDirectory", Path.Combine(TestDir, "output") }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithoutOutputDirectory_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidSlideRange_ThrowsArgumentException()
    {
        var outputDir = Path.Combine(TestDir, "split_invalid");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "inputPath", _inputPath },
            { "outputDirectory", outputDir },
            { "startSlideIndex", 5 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
