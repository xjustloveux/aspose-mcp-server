using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class AlignPptShapesHandlerTests : PptHandlerTestBase
{
    private readonly AlignPptShapesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Align()
    {
        Assert.Equal("align", _handler.Operation);
    }

    #endregion

    #region Align Left

    [Fact]
    public void Execute_AlignLeft_AlignsToLeftmost()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0, 1 } },
            { "align", "left" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(50, pres.Slides[0].Shapes[0].X);
        Assert.Equal(50, pres.Slides[0].Shapes[1].X);
    }

    #endregion

    #region Align to Slide

    [Fact]
    public void Execute_WithAlignToSlideTrue_AlignsToSlide()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0, 1 } },
            { "align", "left" },
            { "alignToSlide", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(0, pres.Slides[0].Shapes[0].X);
        Assert.Equal(0, pres.Slides[0].Shapes[1].X);
    }

    #endregion

    #region Multiple Shapes

    [Fact]
    public void Execute_WithThreeShapes_AlignsAll()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Triangle, 300, 300, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0, 1, 2 } },
            { "align", "top" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("3", result);
        AssertModified(context);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsShapeCountAndAlignmentInMessage()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0, 1 } },
            { "align", "center" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("2", result);
        Assert.Contains("center", result);
    }

    #endregion

    #region Basic Align Operations

    [Fact]
    public void Execute_AlignsShapes()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0, 1 } },
            { "align", "left" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Aligned", result);
        Assert.Contains("2", result);
        AssertModified(context);
    }

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("top")]
    [InlineData("middle")]
    [InlineData("bottom")]
    public void Execute_SupportsVariousAlignments(string align)
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0, 1 } },
            { "align", align }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains(align, result);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0, 1 } },
            { "align", "left" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutShapeIndices_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "align", "left" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndices", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutAlign_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0, 1 } }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("align", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithLessThanTwoShapes_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0 } },
            { "align", "left" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("2", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidAlign_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0, 1 } },
            { "align", "invalid" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("align", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", new[] { 0, 10 } },
            { "align", "left" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
