using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class ReorderPptShapeHandlerTests : PptHandlerTestBase
{
    private readonly ReorderPptShapeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Reorder()
    {
        Assert.Equal("reorder", _handler.Operation);
    }

    #endregion

    #region Preserve Shape Count

    [Fact]
    public void Execute_PreservesShapeCount()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var initialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "toIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, pres.Slides[0].Shapes.Count);
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_ReordersOnSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 },
            { "toIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("0", result.Message);
        Assert.Contains("1", result.Message);
        AssertModified(context);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsIndicesInMessage()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "toIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("0", result.Message);
        Assert.Contains("1", result.Message);
    }

    #endregion

    #region Basic Reorder Operations

    [Fact]
    public void Execute_ReordersShape()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Triangle, 300, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "toIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("Z-order", result.Message, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(0, 2)]
    [InlineData(2, 0)]
    [InlineData(1, 0)]
    public void Execute_ReordersToVariousPositions(int from, int to)
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Triangle, 300, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", from },
            { "toIndex", to }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains(from.ToString(), result.Message);
        Assert.Contains(to.ToString(), result.Message);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "toIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "toIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutToIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("toIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", invalidIndex },
            { "toIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidToIndex_ThrowsArgumentException(int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "toIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
