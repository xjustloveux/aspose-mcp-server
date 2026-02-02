using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class GroupPptShapesHandlerTests : PptHandlerTestBase
{
    private static readonly int[] TwoShapeIndices = [0, 1];
    private static readonly int[] ThreeShapeIndices = [0, 1, 2];
    private static readonly int[] SingleShapeIndex = [0];
    private static readonly int[] InvalidShapeIndices = [0, 10];

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
            { "shapeIndices", ThreeShapeIndices }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Single(pres.Slides[0].Shapes);
        var groupShape = Assert.IsType<GroupShape>(pres.Slides[0].Shapes[0]);
        Assert.Equal(3, groupShape.Shapes.Count);
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
            { "shapeIndices", TwoShapeIndices }
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
            { "shapeIndices", TwoShapeIndices }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Single(pres.Slides[0].Shapes);
        Assert.IsType<GroupShape>(pres.Slides[0].Shapes[0]);
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
            { "shapeIndices", TwoShapeIndices }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Single(pres.Slides[0].Shapes);
        var groupShape = Assert.IsType<GroupShape>(pres.Slides[0].Shapes[0]);
        Assert.Equal(2, groupShape.Shapes.Count);
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
            { "shapeIndices", TwoShapeIndices }
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
            { "shapeIndices", TwoShapeIndices }
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
        var shape1 = pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        var shape2 = pres.Slides[1].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var idx1 = pres.Slides[1].Shapes.IndexOf(shape1);
        var idx2 = pres.Slides[1].Shapes.IndexOf(shape2);
        var countBefore = pres.Slides[1].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndices", new[] { idx1, idx2 } }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(countBefore - 1, pres.Slides[1].Shapes.Count);
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
            { "shapeIndices", TwoShapeIndices }
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
            { "shapeIndices", SingleShapeIndex }
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
            { "shapeIndices", InvalidShapeIndices }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
