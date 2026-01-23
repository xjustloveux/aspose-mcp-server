using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

public class ClearPptSlideHandlerTests : PptHandlerTestBase
{
    private readonly ClearPptSlideHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Clear()
    {
        Assert.Equal("clear", _handler.Operation);
    }

    #endregion

    #region Multiple Shapes

    [Fact]
    public void Execute_ClearsMultipleShapeTypes()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 100, 10, 50, 50);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Triangle, 200, 10, 50, 50);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Line, 300, 10, 100, 5);
        Assert.True(pres.Slides[0].Shapes.Count >= 4);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Empty(pres.Slides[0].Shapes);
    }

    #endregion

    #region Preserve Other Slides

    [Fact]
    public void Execute_PreservesOtherSlides()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        pres.Slides[2].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 50, 50);
        var slide0Count = pres.Slides[0].Shapes.Count;
        var slide2Count = pres.Slides[2].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(slide0Count, pres.Slides[0].Shapes.Count);
        Assert.Empty(pres.Slides[1].Shapes);
        Assert.Equal(slide2Count, pres.Slides[2].Shapes.Count);
    }

    #endregion

    #region Empty Slide

    [Fact]
    public void Execute_OnEmptySlide_Succeeds()
    {
        var pres = CreatePresentationWithSlides(1);
        while (pres.Slides[0].Shapes.Count > 0)
            pres.Slides[0].Shapes.RemoveAt(0);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Cleared", result.Message);
        Assert.Empty(pres.Slides[0].Shapes);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSlideIndexInMessage()
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

    #region Basic Clear Operations

    [Fact]
    public void Execute_ClearsAllShapesFromSlide()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 300, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Cleared", result.Message);
        Assert.Empty(pres.Slides[0].Shapes);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_ClearsSlideAtVariousIndices(int slideIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[slideIndex].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", slideIndex }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Cleared", result.Message);
        Assert.Empty(pres.Slides[slideIndex].Shapes);
        AssertModified(context);
    }

    #endregion

    #region Error Handling - Invalid Index

    [Theory]
    [InlineData(3, 3)]
    [InlineData(3, 5)]
    [InlineData(3, 100)]
    public void Execute_WithIndexOutOfRange_ThrowsArgumentException(int totalSlides, int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(totalSlides);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", invalidIndex }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(-5)]
    public void Execute_WithNegativeIndex_ThrowsException(int negativeIndex)
    {
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", negativeIndex }
        });

        Assert.ThrowsAny<Exception>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
