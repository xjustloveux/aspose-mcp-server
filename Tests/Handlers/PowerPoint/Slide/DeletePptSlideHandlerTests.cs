using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

public class DeletePptSlideHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptSlideHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Slide Content Preservation

    [Fact]
    public void Execute_PreservesOtherSlidesContent()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[2].Shapes.AddAutoShape(ShapeType.Ellipse, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(2, pres.Slides.Count);
        Assert.True(pres.Slides[0].Shapes.Count > 0, "First slide content should be preserved");
        Assert.True(pres.Slides[1].Shapes.Count > 0, "Last slide content should be preserved");
    }

    #endregion

    #region Error Handling - Missing Parameter

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Delete Operations

    [Fact]
    public void Execute_DeletesSlideFromPresentation()
    {
        var pres = CreatePresentationWithSlides(3);
        var initialCount = pres.Slides.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(initialCount - 1, pres.Slides.Count);
        AssertModified(context);
    }

    [Theory]
    [InlineData(3, 0)]
    [InlineData(3, 1)]
    [InlineData(3, 2)]
    [InlineData(5, 0)]
    [InlineData(5, 4)]
    public void Execute_DeletesSlideAtVariousIndices(int totalSlides, int deleteIndex)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", deleteIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(totalSlides - 1, pres.Slides.Count);
        AssertModified(context);
    }

    [Theory]
    [InlineData(3, 0, 2)]
    [InlineData(3, 2, 2)]
    [InlineData(5, 2, 4)]
    public void Execute_DeletesAtPosition_CorrectCountRemains(int initial, int deleteIndex, int expectedRemaining)
    {
        var pres = CreatePresentationWithSlides(initial);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", deleteIndex }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(expectedRemaining, pres.Slides.Count);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsRemainingSlideCount()
    {
        var pres = CreatePresentationWithSlides(5);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("4", result.Message);
        Assert.Contains("remaining", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_ReturnsDeletedSlideIndex()
    {
        var pres = CreatePresentationWithSlides(5);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 3 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("3", result.Message);
    }

    #endregion

    #region Error Handling - Invalid Index

    [Theory]
    [InlineData(3, 3)]
    [InlineData(3, 5)]
    [InlineData(3, 10)]
    [InlineData(3, 100)]
    public void Execute_WithIndexOutOfRange_ThrowsArgumentException(int totalSlides, int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-5)]
    [InlineData(-100)]
    public void Execute_WithNegativeIndex_ThrowsArgumentException(int negativeIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", negativeIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Error Handling - Last Slide

    [Fact]
    public void Execute_LastRemainingSlide_ThrowsInvalidOperationException()
    {
        var pres = CreateEmptyPresentation();
        Assert.Single(pres.Slides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var ex = Assert.Throws<InvalidOperationException>(() => _handler.Execute(context, parameters));
        Assert.Contains("last slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_CanDeleteUntilOneRemains()
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "slideIndex", 0 } }));
        Assert.Equal(2, pres.Slides.Count);

        _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "slideIndex", 0 } }));
        Assert.Single(pres.Slides);

        var ex = Assert.Throws<InvalidOperationException>(() =>
            _handler.Execute(context, CreateParameters(new Dictionary<string, object?> { { "slideIndex", 0 } })));
        Assert.Contains("last slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
