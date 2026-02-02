using Aspose.Slides.SlideShow;
using AsposeMcpServer.Handlers.PowerPoint.Transition;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Transition;

public class SetPptTransitionHandlerTests : PptHandlerTestBase
{
    private readonly SetPptTransitionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Set()
    {
        Assert.Equal("set", _handler.Operation);
    }

    #endregion

    #region Multiple Slides

    [Fact]
    public void Execute_WithSlideIndex_SetsOnCorrectSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "transitionType", "Wipe" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(TransitionType.None, pres.Slides[0].SlideShowTransition.Type);
        Assert.Equal(TransitionType.Wipe, pres.Slides[1].SlideShowTransition.Type);
    }

    #endregion

    #region Basic Set Operations

    [Fact]
    public void Execute_SetsTransition()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "transitionType", "Fade" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(TransitionType.Fade, pres.Slides[0].SlideShowTransition.Type);
        AssertModified(context);
    }

    [Theory]
    [InlineData("Fade", TransitionType.Fade)]
    [InlineData("Wipe", TransitionType.Wipe)]
    [InlineData("Push", TransitionType.Push)]
    [InlineData("Dissolve", TransitionType.Dissolve)]
    [InlineData("None", TransitionType.None)]
    [InlineData("Blinds", TransitionType.Blinds)]
    [InlineData("Checker", TransitionType.Checker)]
    [InlineData("Circle", TransitionType.Circle)]
    [InlineData("Comb", TransitionType.Comb)]
    [InlineData("Cover", TransitionType.Cover)]
    [InlineData("Cut", TransitionType.Cut)]
    [InlineData("Diamond", TransitionType.Diamond)]
    [InlineData("Newsflash", TransitionType.Newsflash)]
    [InlineData("Plus", TransitionType.Plus)]
    [InlineData("Random", TransitionType.Random)]
    [InlineData("RandomBar", TransitionType.RandomBar)]
    [InlineData("Split", TransitionType.Split)]
    [InlineData("Strips", TransitionType.Strips)]
    [InlineData("Wedge", TransitionType.Wedge)]
    [InlineData("Wheel", TransitionType.Wheel)]
    [InlineData("Zoom", TransitionType.Zoom)]
    public void Execute_WithVariousTypes_SetsCorrectTransition(string typeStr, TransitionType expectedType)
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "transitionType", typeStr }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(expectedType, pres.Slides[0].SlideShowTransition.Type);
    }

    [Fact]
    public void Execute_WithAdvanceAfterSeconds_SetsAdvance()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "transitionType", "Fade" },
            { "advanceAfterSeconds", 3.5 }
        });

        _handler.Execute(context, parameters);

        Assert.True(pres.Slides[0].SlideShowTransition.AdvanceAfter);
        Assert.Equal(3500u, pres.Slides[0].SlideShowTransition.AdvanceAfterTime);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "transitionType", "Fade" },
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(TransitionType.Fade, pres.Slides[0].SlideShowTransition.Type);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutTransitionType_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("transitionType", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidTransitionType_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "transitionType", "InvalidType" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Invalid transition type", ex.Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Execute_WithEmptyTransitionType_ThrowsArgumentException(string emptyType)
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "transitionType", emptyType }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("transitionType is required", ex.Message);
    }

    [Fact]
    public void Execute_WithAdvanceAfterSecondsZero_DisablesAdvance()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "transitionType", "Fade" },
            { "advanceAfterSeconds", 0.0 }
        });

        _handler.Execute(context, parameters);

        Assert.False(pres.Slides[0].SlideShowTransition.AdvanceAfter);
        Assert.Equal(0u, pres.Slides[0].SlideShowTransition.AdvanceAfterTime);
    }

    [Theory]
    [InlineData("fade")]
    [InlineData("FADE")]
    [InlineData("FaDe")]
    public void Execute_WithCaseInsensitiveType_SetsTransition(string typeCase)
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "transitionType", typeCase }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(TransitionType.Fade, pres.Slides[0].SlideShowTransition.Type);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "transitionType", "Fade" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
