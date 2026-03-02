using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Handlers.PowerPoint.Transition;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Transition;

[SupportedOSPlatform("windows")]
public class DeletePptTransitionHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptTransitionHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Delete()
    {
        SkipIfNotWindows();
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Multiple Slides

    [SkippableFact]
    public void Execute_WithSlideIndex_RemovesFromCorrectSlide()
    {
        SkipIfNotWindows();
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

    #region Can Remove Non-Existent Transition

    [SkippableFact]
    public void Execute_WithNoExistingTransition_StillSucceeds()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message, StringComparison.OrdinalIgnoreCase);
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

    [SkippableFact]
    public void Execute_RemovesTransition()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTransition(TransitionType.Fade);
        Assert.Equal(TransitionType.Fade, pres.Slides[0].SlideShowTransition.Type);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("removed", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(TransitionType.None, pres.Slides[0].SlideShowTransition.Type);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_SetsAdvanceOnClick()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithTransition(TransitionType.Fade);
        pres.Slides[0].SlideShowTransition.AdvanceOnClick = false;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.True(pres.Slides[0].SlideShowTransition.AdvanceOnClick);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("slide 0", result.Message);
    }

    #endregion
}
