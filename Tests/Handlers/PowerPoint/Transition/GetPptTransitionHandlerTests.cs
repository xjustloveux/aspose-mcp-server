using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Handlers.PowerPoint.Transition;
using AsposeMcpServer.Results.PowerPoint.Transition;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Transition;

public class GetPptTransitionHandlerTests : PptHandlerTestBase
{
    private readonly GetPptTransitionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Multiple Slides

    [Fact]
    public void Execute_WithSlideIndex_ReturnsCorrectSlideInfo()
    {
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

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
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

    [Fact]
    public void Execute_DoesNotModifyDocument()
    {
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

    [Fact]
    public void Execute_ReturnsTransitionInfo()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.NotNull(result.Type);
        Assert.False(result.HasTransition);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithNoTransition_ReturnsNone()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.Equal("None", result.Type);
        Assert.False(result.HasTransition);
    }

    [Fact]
    public void Execute_WithTransition_ReturnsTransitionType()
    {
        var pres = CreatePresentationWithTransition(TransitionType.Fade);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.Equal("Fade", result.Type);
        Assert.True(result.HasTransition);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.Equal(0, result.SlideIndex);
    }

    [Fact]
    public void Execute_ReturnsAdvanceSettings()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetTransitionResult>(res);

        Assert.NotNull(result);
        Assert.True(result.AdvanceOnClick || !result.AdvanceOnClick);
        Assert.True(result.AdvanceAfter || !result.AdvanceAfter);
        Assert.True(result.AdvanceAfterSeconds >= 0 || result.AdvanceAfterSeconds < 0);
    }

    #endregion
}
