using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Handlers.PowerPoint.Transition;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Transition;

public class DeletePptTransitionHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptTransitionHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Multiple Slides

    [Fact]
    public void Execute_WithSlideIndex_RemovesFromCorrectSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].SlideShowTransition.Type = TransitionType.Fade;
        pres.Slides[1].SlideShowTransition.Type = TransitionType.Wipe;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(TransitionType.Fade, pres.Slides[0].SlideShowTransition.Type);
        Assert.Equal(TransitionType.None, pres.Slides[1].SlideShowTransition.Type);
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

    #region Can Remove Non-Existent Transition

    [Fact]
    public void Execute_WithNoExistingTransition_StillSucceeds()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result.ToLower());
        AssertModified(context);
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

    #region Basic Delete Operations

    [Fact]
    public void Execute_RemovesTransition()
    {
        var pres = CreatePresentationWithTransition(TransitionType.Fade);
        Assert.Equal(TransitionType.Fade, pres.Slides[0].SlideShowTransition.Type);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("removed", result.ToLower());
        Assert.Equal(TransitionType.None, pres.Slides[0].SlideShowTransition.Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_SetsAdvanceOnClick()
    {
        var pres = CreatePresentationWithTransition(TransitionType.Fade);
        pres.Slides[0].SlideShowTransition.AdvanceOnClick = false;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.True(pres.Slides[0].SlideShowTransition.AdvanceOnClick);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide 0", result);
    }

    #endregion
}
