using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.PowerPoint.Image;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Image;

[SupportedOSPlatform("windows")]
[Collection("SerialSlides")]
public class ExportSlidesHandlerTests : PptHandlerTestBase
{
    private readonly ExportSlidesHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_ExportSlides()
    {
        SkipIfNotWindows();
        Assert.Equal("export_slides", _handler.Operation);
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

    /// <summary>
    ///     §18.4.3: the output directory was created before it was validated, so a request for a
    ///     path outside the allowlist left a directory behind even though it was refused.
    /// </summary>
    [Fact]
    public void Execute_WithAnOutputDirOutsideTheAllowlist_ShouldNotCreateIt()
    {
        var outside = Path.Combine(Path.GetTempPath(), "ExportSlidesOutside_" + Guid.NewGuid().ToString("N"));
        var presentation = CreatePresentationWithSlides(2);

        // AllowedBasePaths has a private setter, so the config is built the way the other
        // security-focused suites build one, and the context is created around it.
        var config = new ServerConfig();
        typeof(ServerConfig)
            .GetProperty(nameof(ServerConfig.AllowedBasePaths))!
            .SetValue(config, new List<string> { Path.GetFullPath(TestDir) }.AsReadOnly());

        var context = new OperationContext<Presentation>
        {
            Document = presentation,
            ServerConfig = config
        };

        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputDir", outside }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.False(Directory.Exists(outside), "the refused request created its output directory");
    }

    #region Basic Export Operations

    [SkippableFact]
    public void Execute_ExportsSlidesToImages()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithSpecificSlideIndexes_ExportsSelectedSlides()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithJpegFormat_ExportsAsJpeg()
    {
        SkipIfNotWindows();
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
