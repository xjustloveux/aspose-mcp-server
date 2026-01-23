using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.PowerPoint.Shape;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class GetPptShapeDetailsHandlerTests : PptHandlerTestBase
{
    private readonly GetPptShapeDetailsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetShapeDetails()
    {
        Assert.Equal("get_shape_details", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyPresentation()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var initialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, pres.Slides[0].Shapes.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsShapeDetails()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 300, 400);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.True(result.Index >= 0);
        Assert.NotNull(result.Type);
        Assert.True(result.X >= 0);
        Assert.True(result.Y >= 0);
        Assert.True(result.Width >= 0);
        Assert.True(result.Height >= 0);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectPosition()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 150, 250, 300, 400);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(150f, result.X);
        Assert.Equal(250f, result.Y);
    }

    [Fact]
    public void Execute_ReturnsCorrectDimensions()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 350, 450);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(350f, result.Width);
        Assert.Equal(450f, result.Height);
    }

    #endregion

    #region AutoShape Details

    [Fact]
    public void Execute_ForAutoShape_ReturnsShapeType()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.NotNull(result.ShapeType);
        Assert.Equal("Ellipse", result.ShapeType);
    }

    [Fact]
    public void Execute_ForAutoShape_ReturnsFillType()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.FillFormat.FillType = FillType.Solid;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.NotNull(result.FillType);
        Assert.Equal("Solid", result.FillType);
    }

    #endregion

    #region Rotation and Hidden

    [Fact]
    public void Execute_ReturnsRotation()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.Rotation = 45;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(45f, result.Rotation);
    }

    [Fact]
    public void Execute_ReturnsHiddenStatus()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.Hidden = true;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.True(result.Hidden);
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_GetsFromSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Triangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(0, result.Index);
    }

    [Fact]
    public void Execute_DefaultSlideIndex_GetsFromFirstSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);
        var result = Assert.IsType<GetShapeDetailsResult>(res);

        Assert.Equal(0, result.Index);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
