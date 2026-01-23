using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Results.PowerPoint.Slide;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

public class GetPptSlidesInfoHandlerTests : PptHandlerTestBase
{
    private readonly GetPptSlidesInfoHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetInfo()
    {
        Assert.Equal("get_info", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyPresentation()
    {
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

    [Fact]
    public void Execute_ReturnsSlideInfo()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        Assert.Equal(3, result.Count);
        AssertNotModified(context);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(5)]
    [InlineData(10)]
    public void Execute_ReturnsCorrectSlideCount(int slideCount)
    {
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

    [Fact]
    public void Execute_ReturnsSlidesArray()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        Assert.NotNull(result.Slides);
        Assert.Equal(3, result.Slides.Count);
    }

    [Fact]
    public void Execute_SlidesContainIndex()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        var firstSlide = result.Slides[0];
        Assert.Equal(0, firstSlide.Index);
    }

    [Fact]
    public void Execute_SlidesContainLayoutInfo()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        var firstSlide = result.Slides[0];
        Assert.NotNull(firstSlide.LayoutType);
    }

    [Fact]
    public void Execute_SlidesContainShapesCount()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        var firstSlide = result.Slides[0];
        Assert.True(firstSlide.ShapesCount >= 1);
    }

    [Fact]
    public void Execute_SlidesContainHiddenProperty()
    {
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

    [Fact]
    public void Execute_ReturnsAvailableLayouts()
    {
        var pres = CreatePresentationWithSlides(1);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlidesInfoResult>(res);

        Assert.NotNull(result.AvailableLayouts);
        Assert.True(result.AvailableLayouts.Count > 0);
    }

    [Fact]
    public void Execute_LayoutsContainNameAndType()
    {
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
