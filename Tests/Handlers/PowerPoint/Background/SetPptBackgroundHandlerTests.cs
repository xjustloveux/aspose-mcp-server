using System.Drawing;
using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Background;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Background;

[SupportedOSPlatform("windows")]
public class SetPptBackgroundHandlerTests : PptHandlerTestBase
{
    private readonly SetPptBackgroundHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Set()
    {
        SkipIfNotWindows();
        Assert.Equal("set", _handler.Operation);
    }

    #endregion

    #region Basic Set Operations

    [SkippableFact]
    public void Execute_SetsBackgroundColor()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "color", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var slide = pres.Slides[0];
        Assert.Equal(BackgroundType.OwnBackground, slide.Background.Type);
        Assert.Equal(FillType.Solid, slide.Background.FillFormat.FillType);
        if (!IsEvaluationMode())
            Assert.Equal(Color.FromArgb(255, 0, 0), slide.Background.FillFormat.SolidFillColor.Color);

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithSlideIndex_SetsBackgroundOnSpecificSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "color", "#00FF00" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        var slide = pres.Slides[1];
        Assert.Equal(BackgroundType.OwnBackground, slide.Background.Type);
        Assert.Equal(FillType.Solid, slide.Background.FillFormat.FillType);
        if (!IsEvaluationMode())
            Assert.Equal(Color.FromArgb(0, 255, 0), slide.Background.FillFormat.SolidFillColor.Color);

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithApplyToAll_SetsBackgroundOnAllSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "color", "#0000FF" },
            { "applyToAll", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        foreach (var slide in pres.Slides)
        {
            Assert.Equal(BackgroundType.OwnBackground, slide.Background.Type);
            Assert.Equal(FillType.Solid, slide.Background.FillFormat.FillType);
            if (!IsEvaluationMode())
                Assert.Equal(Color.FromArgb(0, 0, 255), slide.Background.FillFormat.SolidFillColor.Color);
        }

        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutColorOrImage_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "color", "#FF0000" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
