using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class GroupPptShapesHandlerTests : PptHandlerTestBase
{
    private readonly GroupPptShapesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Group()
    {
        Assert.Equal("group", _handler.Operation);
    }

    #endregion

    #region Multiple Shapes

    [Fact]
    public void Execute_WithThreeShapes_GroupsAll()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Triangle, 300, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0, 1, 2 } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("3", result);
        Assert.Single(pres.Slides[0].Shapes);
        AssertModified(context);
    }

    #endregion

    #region Group Position and Size

    [Fact]
    public void Execute_GroupIsCreated()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0, 1 } }
        });

        _handler.Execute(context, parameters);

        var groupShape = pres.Slides[0].Shapes.OfType<IGroupShape>().FirstOrDefault();
        Assert.NotNull(groupShape);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsGroupIndexInMessage()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0, 1 } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("shapeIndex", result);
    }

    #endregion

    #region Basic Group Operations

    [Fact]
    public void Execute_GroupsShapes()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0, 1 } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Grouped", result);
        Assert.Contains("2", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_CreatesGroupShape()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0, 1 } }
        });

        _handler.Execute(context, parameters);

        Assert.Contains(pres.Slides[0].Shapes.OfType<IGroupShape>(), s => s != null);
    }

    [Fact]
    public void Execute_RemovesOriginalShapes()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0, 1 } }
        });

        _handler.Execute(context, parameters);

        Assert.Single(pres.Slides[0].Shapes);
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_GroupsOnSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndices", new[] { 0, 1 } }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Grouped", result);
        Assert.Contains(pres.Slides[1].Shapes.OfType<IGroupShape>(), s => s != null);
    }

    [Fact]
    public void Execute_DefaultSlideIndex_GroupsOnFirstSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0, 1 } }
        });

        _handler.Execute(context, parameters);

        Assert.Contains(pres.Slides[0].Shapes.OfType<IGroupShape>(), s => s != null);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndices_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndices", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithLessThanTwoShapes_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0 } }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("2", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", new[] { 0, 10 } }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
