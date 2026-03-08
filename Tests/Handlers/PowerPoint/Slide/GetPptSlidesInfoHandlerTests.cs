using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Results.PowerPoint.Slide;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

[SupportedOSPlatform("windows")]
public class GetPptSlidesInfoHandlerTests : PptHandlerTestBase
{
    private readonly GetPptSlidesInfoHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetInfo()
    {
        SkipIfNotWindows();
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [SkippableFact]
    public void Execute_DoesNotModifyPresentation()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, pres.Slides.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Basic Info Retrieval

    [SkippableFact]
    public void Execute_ReturnsSlideInfo()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        Assert.Equal(3, result.Count);
        AssertNotModified(context);
    }

    [SkippableTheory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(5)]
    [InlineData(10)]
    public void Execute_ReturnsCorrectSlideCount(int slideCount)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(slideCount);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        Assert.Equal(slideCount, result.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Slides Array

    [SkippableFact]
    public void Execute_ReturnsSlidesArray()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        Assert.NotNull(result.Slides);
        Assert.Equal(3, result.Slides.Count);
    }

    [SkippableFact]
    public void Execute_SlidesContainIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        var firstSlide = result.Slides[0];
        Assert.Equal(0, firstSlide.Index);
    }

    [SkippableFact]
    public void Execute_SlidesContainLayoutInfo()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        var firstSlide = result.Slides[0];
        Assert.NotNull(firstSlide.LayoutType);
    }

    [SkippableFact]
    public void Execute_SlidesContainShapesCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        var firstSlide = result.Slides[0];
        Assert.True(firstSlide.ShapesCount >= 1);
    }

    [SkippableFact]
    public void Execute_SlidesContainHiddenProperty()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        var firstSlide = result.Slides[0];
        Assert.False(firstSlide.Hidden);
    }

    #endregion

    #region Available Layouts

    [SkippableFact]
    public void Execute_ReturnsAvailableLayouts()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        Assert.NotNull(result.AvailableLayouts);
        Assert.True(result.AvailableLayouts.Count > 0);
    }

    [SkippableFact]
    public void Execute_LayoutsContainNameAndType()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        var firstLayout = result.AvailableLayouts[0];
        Assert.True(firstLayout.Index >= 0);
        Assert.NotNull(firstLayout.Type);
    }

    #endregion
}
