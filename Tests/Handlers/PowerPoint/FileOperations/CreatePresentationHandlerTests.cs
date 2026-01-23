using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.FileOperations;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.FileOperations;

public class CreatePresentationHandlerTests : PptHandlerTestBase
{
    private readonly CreatePresentationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Create()
    {
        Assert.Equal("create", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPathOrOutputPath_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Create Operations

    [Fact]
    public void Execute_WithPath_CreatesPresentation()
    {
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

    [Fact]
    public void Execute_WithOutputPath_CreatesPresentation()
    {
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

    [Fact]
    public void Execute_CreatesValidPresentation()
    {
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
