using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class UngroupPptShapesHandlerTests : PptHandlerTestBase
{
    private readonly UngroupPptShapesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Ungroup()
    {
        Assert.Equal("ungroup", _handler.Operation);
    }

    #endregion

    #region Group with Three Shapes

    [Fact]
    public void Execute_WithThreeShapesInGroup_UngroupsAll()
    {
        var pres = CreatePresentationWithSlides(1);
        var groupShape = pres.Slides[0].Shapes.AddGroupShape();
        groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Triangle, 200, 0, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("3", result);
        Assert.Equal(3, pres.Slides[0].Shapes.Count);
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_UngroupsOnSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var groupShape = pres.Slides[1].Shapes.AddGroupShape();
        groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 0, 100, 100);
        var groupIndex = pres.Slides[1].Shapes.IndexOf(groupShape);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", groupIndex }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide 1", result);
        Assert.DoesNotContain(pres.Slides[1].Shapes.OfType<IGroupShape>(), s => s != null);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsShapeCountInMessage()
    {
        var pres = CreatePresentationWithSlides(1);
        var groupShape = pres.Slides[0].Shapes.AddGroupShape();
        groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 0, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("2", result);
    }

    #endregion

    #region Basic Ungroup Operations

    [Fact]
    public void Execute_UngroupsShapes()
    {
        var pres = CreatePresentationWithSlides(1);
        var groupShape = pres.Slides[0].Shapes.AddGroupShape();
        groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 0, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Ungrouped", result);
        Assert.Contains("2", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_RemovesGroupShape()
    {
        var pres = CreatePresentationWithSlides(1);
        var groupShape = pres.Slides[0].Shapes.AddGroupShape();
        groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 0, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.DoesNotContain(pres.Slides[0].Shapes.OfType<IGroupShape>(), s => s != null);
    }

    [Fact]
    public void Execute_CreatesIndividualShapes()
    {
        var pres = CreatePresentationWithSlides(1);
        var groupShape = pres.Slides[0].Shapes.AddGroupShape();
        groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 0, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(2, pres.Slides[0].Shapes.Count);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        var groupShape = pres.Slides[0].Shapes.AddGroupShape();
        groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 0, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        var groupShape = pres.Slides[0].Shapes.AddGroupShape();
        groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 0, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNonGroupShape_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("group", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(1);
        var groupShape = pres.Slides[0].Shapes.AddGroupShape();
        groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 100, 100);
        groupShape.Shapes.AddAutoShape(ShapeType.Ellipse, 100, 0, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
