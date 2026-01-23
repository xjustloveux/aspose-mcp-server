using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.PageSetup;

public class SetSlideOrientationHandlerTests : PptHandlerTestBase
{
    private readonly SetSlideOrientationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetOrientation()
    {
        Assert.Equal("set_orientation", _handler.Operation);
    }

    #endregion

    #region Basic Set Slide Orientation Operations

    [Fact]
    public void Execute_SetsPortraitOrientation()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "Portrait" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("portrait", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsLandscapeOrientation()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "Landscape" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("landscape", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_ReturnsSizeInfo()
    {
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "orientation", "Portrait" }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("x", result.Message);
    }

    #endregion
}
