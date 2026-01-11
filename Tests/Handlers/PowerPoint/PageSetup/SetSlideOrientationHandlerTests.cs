using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Tests.Helpers;

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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("portrait", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("landscape", result.ToLower());
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

        var result = _handler.Execute(context, parameters);

        Assert.Contains("x", result);
    }

    #endregion
}
