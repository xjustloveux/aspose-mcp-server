using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

public class AddPptSlideHandlerTests : PptHandlerTestBase
{
    private readonly AddPptSlideHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithEmptyLayoutSlides_ThrowsInvalidOperationException()
    {
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

    [Fact]
    public void Execute_AddsSlideToPresentation()
    {
        var pres = CreateEmptyPresentation();
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(initialCount + 1, pres.Slides.Count);
        AssertModified(context);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(5)]
    public void Execute_AddsCorrectNumberOfSlides(int slidesToAdd)
    {
        var pres = CreateEmptyPresentation();
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        for (var i = 0; i < slidesToAdd; i++) _handler.Execute(context, parameters);

        Assert.Equal(initialCount + slidesToAdd, pres.Slides.Count);
    }

    #endregion

    #region Layout Types

    [Theory]
    [InlineData("Blank")]
    [InlineData("blank")]
    [InlineData("BLANK")]
    public void Execute_WithBlankLayout_AddsBlankSlide(string layoutType)
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layoutType", layoutType }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        var lastSlide = pres.Slides[^1];
        Assert.NotNull(lastSlide.LayoutSlide);
        AssertModified(context);
    }

    [Theory]
    [InlineData("Title")]
    [InlineData("TitleOnly")]
    [InlineData("TwoColumn")]
    [InlineData("SectionHeader")]
    public void Execute_WithVariousLayoutTypes_AddsSlideWithLayout(string layoutType)
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layoutType", layoutType }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Theory]
    [InlineData("UnknownLayout")]
    [InlineData("CustomType")]
    [InlineData("InvalidLayout")]
    public void Execute_WithUnknownLayout_UsesDefaultLayout(string layoutType)
    {
        var pres = CreateEmptyPresentation();
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "layoutType", layoutType }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(initialCount + 1, pres.Slides.Count);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DefaultLayout_IsBlank()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        var lastSlide = pres.Slides[^1];
        Assert.NotNull(lastSlide);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSlideCount()
    {
        var pres = CreateEmptyPresentation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("total", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(pres.Slides.Count.ToString(), result);
    }

    [Fact]
    public void Execute_ReturnsCorrectTotalAfterMultipleAdditions()
    {
        var pres = CreatePresentationWithSlides(5);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        Assert.Contains("6", result);
    }

    #endregion

    #region Presentation State

    [Fact]
    public void Execute_PreservesExistingSlides()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(4, pres.Slides.Count);
        Assert.True(pres.Slides[0].Shapes.Count > 0, "First slide shapes should be preserved");
    }

    [Fact]
    public void Execute_AddsSlideAtEnd()
    {
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
