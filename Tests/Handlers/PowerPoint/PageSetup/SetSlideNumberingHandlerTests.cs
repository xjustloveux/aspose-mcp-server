using System.Runtime.Versioning;
using AsposeMcpServer.Handlers.PowerPoint.PageSetup;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.PageSetup;

[SupportedOSPlatform("windows")]
public class SetSlideNumberingHandlerTests : PptHandlerTestBase
{
    private readonly SetSlideNumberingHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_SetSlideNumbering()
    {
        SkipIfNotWindows();
        Assert.Equal("set_slide_numbering", _handler.Operation);
    }

    #endregion

    #region Basic Set Slide Numbering Operations

    [SkippableFact]
    public void Execute_ShowsSlideNumbers()
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
        AssertModified(context);
        if (!IsEvaluationMode())
        {
            Assert.Equal(1, presentation.FirstSlideNumber);
            foreach (var slide in presentation.Slides)
                Assert.True(slide.HeaderFooterManager.IsSlideNumberVisible);
        }
    }

    [SkippableFact]
    public void Execute_HidesSlideNumbers()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "showSlideNumber", false }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            Assert.Equal(1, presentation.FirstSlideNumber);
            foreach (var slide in presentation.Slides)
                Assert.False(slide.HeaderFooterManager.IsSlideNumberVisible);
        }
    }

    [SkippableFact]
    public void Execute_SetsFirstNumber()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "firstNumber", 5 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(5, presentation.FirstSlideNumber);
    }

    [SkippableFact]
    public void Execute_WithDefaults()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(1, presentation.FirstSlideNumber);
        if (!IsEvaluationMode())
            foreach (var slide in presentation.Slides)
                Assert.True(slide.HeaderFooterManager.IsSlideNumberVisible);
    }

    #endregion
}
