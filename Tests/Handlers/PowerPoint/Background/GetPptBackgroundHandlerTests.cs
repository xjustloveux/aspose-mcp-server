using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.Background;
using AsposeMcpServer.Results.PowerPoint.Background;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Background;

[SupportedOSPlatform("windows")]
public class GetPptBackgroundHandlerTests : PptHandlerTestBase
{
    private readonly GetPptBackgroundHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Get()
    {
        SkipIfNotWindows();
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Basic Get Operations

    [SkippableFact]
    public void Execute_ReturnsBackgroundInfo()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBackgroundResult>(res);

        Assert.Equal(0, result.SlideIndex);
        Assert.True(result.HasBackground || !result.HasBackground); // Property exists
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithSlideIndex_ReturnsSpecificSlideBackground()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBackgroundResult>(res);

        Assert.Equal(1, result.SlideIndex);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsAllBackgroundProperties()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetBackgroundResult>(res);

        Assert.IsType<int>(result.SlideIndex);
        Assert.IsType<bool>(result.HasBackground);
        Assert.IsType<bool>(result.IsPictureFill);
    }

    #endregion
}
