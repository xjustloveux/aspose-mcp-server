using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.FileOperations;

[SupportedOSPlatform("windows")]
public class CreatePresentationHandlerTests : PptHandlerTestBase
{
    private readonly CreatePresentationHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Create()
    {
        SkipIfNotWindows();
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutPathOrOutputPath_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Create Operations

    [SkippableFact]
    public void Execute_WithPath_CreatesPresentation()
    {
        SkipIfNotWindows();
        var outputPath = Path.Combine(TestDir, "test.pptx");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Created presentation should have content");
    }

    [SkippableFact]
    public void Execute_WithOutputPath_CreatesPresentation()
    {
        SkipIfNotWindows();
        var outputPath = Path.Combine(TestDir, "output.pptx");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", outputPath }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("created successfully", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(File.Exists(outputPath));
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Created presentation should have content");
    }

    [SkippableFact]
    public void Execute_CreatesValidPresentation()
    {
        SkipIfNotWindows();
        var outputPath = Path.Combine(TestDir, "valid.pptx");
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", outputPath }
        });

        _handler.Execute(context, parameters);

        using var createdPres = new Presentation(outputPath);
        Assert.NotNull(createdPres);
        Assert.Single(createdPres.Slides);
    }

    #endregion
}
