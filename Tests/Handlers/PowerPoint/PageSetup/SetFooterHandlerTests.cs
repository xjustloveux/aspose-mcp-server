using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.PageSetup;

[SupportedOSPlatform("windows")]
public class SetFooterHandlerTests : PptHandlerTestBase
{
    private static readonly int[] SlideIndicesZeroTwo = [0, 2];

    private readonly SetFooterHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_SetFooter()
    {
        SkipIfNotWindows();
        Assert.Equal("set_footer", _handler.Operation);
    }

    #endregion

    #region Basic Set Footer Operations

    [SkippableFact]
    public void Execute_SetsFooterText()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "My Footer" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
            Assert.True(presentation.Slides[0].HeaderFooterManager.IsFooterVisible);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_SetsDateText()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "dateText", "2026-01-11" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
            Assert.True(presentation.Slides[0].HeaderFooterManager.IsDateTimeVisible);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_ShowsSlideNumber()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showSlideNumber", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
            Assert.True(presentation.Slides[0].HeaderFooterManager.IsSlideNumberVisible);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_AppliesToAllSlides()
    {
        SkipIfNotWindows();
        var presentation = CreatePresentationWithSlides(3);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "Test Footer" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
            foreach (var slide in presentation.Slides)
                Assert.True(slide.HeaderFooterManager.IsFooterVisible,
                    $"Footer should be visible on slide {presentation.Slides.IndexOf(slide)}");
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_AppliesToSpecificSlides()
    {
        SkipIfNotWindows();
        var presentation = CreatePresentationWithSlides(3);
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "footerText", "Test Footer" },
            { "slideIndices", SlideIndicesZeroTwo }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            Assert.True(presentation.Slides[0].HeaderFooterManager.IsFooterVisible,
                "Footer should be visible on slide 0");
            Assert.True(presentation.Slides[2].HeaderFooterManager.IsFooterVisible,
                "Footer should be visible on slide 2");
        }

        AssertModified(context);
    }

    #endregion
}
