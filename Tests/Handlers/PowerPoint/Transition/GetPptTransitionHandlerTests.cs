using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Handlers.PowerPoint.Transition;
using AsposeMcpServer.Results.PowerPoint.Transition;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Transition;

[SupportedOSPlatform("windows")]
public class GetPptTransitionHandlerTests : PptHandlerTestBase
{
    private readonly GetPptTransitionHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Get()
    {
        SkipIfNotWindows();
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Multiple Slides

    [SkippableFact]
    public void Execute_WithSlideIndex_ReturnsCorrectSlideInfo()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].SlideShowTransition.Type = TransitionType.Wipe;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.Equal(1, result.SlideIndex);
        Assert.Equal("Wipe", result.Type);
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

    #region Read-Only Verification

    [SkippableFact]
    public void Execute_DoesNotModifyDocument()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        AssertNotModified(context);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithTransition(TransitionType type)
    {
        var pres = new Presentation();
        pres.Slides[0].SlideShowTransition.Type = type;
        return pres;
    }

    #endregion

    #region Basic Get Operations

    [SkippableFact]
    public void Execute_ReturnsTransitionInfo()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.NotNull(result.Type);
        Assert.False(result.HasTransition);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithNoTransition_ReturnsNone()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.Equal("None", result.Type);
        Assert.False(result.HasTransition);
    }

    [SkippableFact]
    public void Execute_WithTransition_ReturnsTransitionType()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTransition(TransitionType.Fade);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.Equal("Fade", result.Type);
        Assert.True(result.HasTransition);
        Assert.Equal(0, result.SlideIndex);
        Assert.NotNull(result.Speed);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.Equal(0, result.SlideIndex);
    }

    [SkippableFact]
    public void Execute_ReturnsAdvanceSettings()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.NotNull(result);
        Assert.Equal(0, result.SlideIndex);
        Assert.Equal("None", result.Type);
        Assert.False(result.HasTransition);
        Assert.False(result.AdvanceAfter);
        Assert.Equal(0.0, result.AdvanceAfterSeconds);
    }

    #endregion
}
