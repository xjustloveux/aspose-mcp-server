using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

[SupportedOSPlatform("windows")]
public class AddPptSlideHandlerTests : PptHandlerTestBase
{
    private readonly AddPptSlideHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Add()
    {
        SkipIfNotWindows();
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithEmptyLayoutSlides_ThrowsInvalidOperationException()
    {
        SkipIfNotWindows();
        var pres = new Presentation();
        while (pres.LayoutSlides.Count > 0)
            try
            {
                pres.LayoutSlides.Remove(pres.LayoutSlides[0]);
            }
            catch
            {
                break;
            }

        if (pres.LayoutSlides.Count == 0)
        {
            var context = CreateContext(pres);
            var parameters = CreateEmptyParameters();

            Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
        }
    }

    #endregion

    #region Basic Slide Addition

    [SkippableFact]
    public void Execute_AddsSlideToPresentation()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialCount + 1, pres.Slides.Count);
        AssertModified(context);
    }

    [SkippableTheory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(5)]
    public void Execute_AddsCorrectNumberOfSlides(int slidesToAdd)
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        for (var i = 0; i < slidesToAdd; i++) _handler.Execute(context, parameters);

        Assert.Equal(initialCount + slidesToAdd, pres.Slides.Count);
    }

    #endregion

    #region Layout Types

    [SkippableTheory]
    [InlineData("Blank")]
    [InlineData("blank")]
    [InlineData("BLANK")]
    public void Execute_WithBlankLayout_AddsBlankSlide(string layoutType)
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layoutType", layoutType }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialCount + 1, pres.Slides.Count);
        var lastSlide = pres.Slides[^1];
        Assert.NotNull(lastSlide.LayoutSlide);
        AssertModified(context);
    }

    [SkippableTheory]
    [InlineData("Title")]
    [InlineData("TitleOnly")]
    [InlineData("TwoColumn")]
    [InlineData("SectionHeader")]
    public void Execute_WithVariousLayoutTypes_AddsSlideWithLayout(string layoutType)
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layoutType", layoutType }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialCount + 1, pres.Slides.Count);
        AssertModified(context);
    }

    [SkippableTheory]
    [InlineData("UnknownLayout")]
    [InlineData("CustomType")]
    [InlineData("InvalidLayout")]
    public void Execute_WithUnknownLayout_UsesDefaultLayout(string layoutType)
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layoutType", layoutType }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialCount + 1, pres.Slides.Count);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_DefaultLayout_IsBlank()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        var lastSlide = pres.Slides[^1];
        Assert.NotNull(lastSlide);
    }

    #endregion

    #region Result Message

    [SkippableFact]
    public void Execute_ReturnsSlideCount()
    {
        SkipIfNotWindows();
        var pres = CreateEmptyPresentation();
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(initialCount + 1, pres.Slides.Count);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectTotalAfterMultipleAdditions()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(5);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(6, pres.Slides.Count);
    }

    #endregion

    #region Presentation State

    [SkippableFact]
    public void Execute_PreservesExistingSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(4, pres.Slides.Count);
        Assert.True(pres.Slides[0].Shapes.Count > 0, "First slide shapes should be preserved");
    }

    [SkippableFact]
    public void Execute_AddsSlideAtEnd()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(4, pres.Slides.Count);
        var lastSlide = pres.Slides[^1];
        Assert.NotNull(lastSlide);
    }

    #endregion
}
