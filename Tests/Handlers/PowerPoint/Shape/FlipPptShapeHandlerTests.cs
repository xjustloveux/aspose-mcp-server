using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class FlipPptShapeHandlerTests : PptHandlerTestBase
{
    private readonly FlipPptShapeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Flip()
    {
        Assert.Equal("flip", _handler.Operation);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsShapeIndexInMessage()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "flipHorizontal", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("0", result.Message);
    }

    #endregion

    #region Basic Flip Operations

    [Fact]
    public void Execute_FlipsHorizontally()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "flipHorizontal", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("flipped", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("H=True", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_FlipsVertically()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "flipVertical", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("flipped", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("V=True", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_FlipsBothDirections()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "flipHorizontal", true },
            { "flipVertical", true }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("H=True", result.Message);
        Assert.Contains("V=True", result.Message);
        AssertModified(context);
    }

    #endregion

    #region Flip State

    [Fact]
    public void Execute_WithFlipHorizontalFalse_SetsNoHorizontalFlip()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "flipHorizontal", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("H=False", result.Message);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithFlipVerticalFalse_SetsNoVerticalFlip()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "flipVertical", false }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("V=False", result.Message);
        AssertModified(context);
    }

    #endregion

    #region Preserve Properties

    [Fact]
    public void Execute_PreservesShapePosition()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 150, 250, 200, 100);
        var originalX = shape.X;
        var originalY = shape.Y;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "flipHorizontal", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(originalX, shape.X);
        Assert.Equal(originalY, shape.Y);
    }

    [Fact]
    public void Execute_PreservesShapeSize()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 250, 150);
        var originalWidth = shape.Width;
        var originalHeight = shape.Height;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "flipVertical", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(originalWidth, shape.Width);
        Assert.Equal(originalHeight, shape.Height);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "flipHorizontal", true }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "flipHorizontal", true }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutFlipOption_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("flip", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", invalidIndex },
            { "flipHorizontal", true }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
