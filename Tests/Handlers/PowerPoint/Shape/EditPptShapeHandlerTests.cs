using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class EditPptShapeHandlerTests : PptHandlerTestBase
{
    private readonly EditPptShapeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Basic Edit Operations

    [Fact]
    public void Execute_UpdatesShape()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "x", 300.0f },
            { "y", 400.0f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(300, pres.Slides[0].Shapes[0].X);
        Assert.Equal(400, pres.Slides[0].Shapes[0].Y);
    }

    #endregion

    #region Rotation

    [Fact]
    public void Execute_WithRotation_UpdatesRotation()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "rotation", 45.0f }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(45, pres.Slides[0].Shapes[0].Rotation);
    }

    #endregion

    #region Name

    [Fact]
    public void Execute_WithName_UpdatesName()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "name", "MyShape" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("MyShape", pres.Slides[0].Shapes[0].Name);
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_EditsOnSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 },
            { "x", 500.0f }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(500, pres.Slides[1].Shapes[0].X);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_WithMultipleProperties_UpdatesAll()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "x", 50.0f },
            { "y", 60.0f },
            { "width", 400.0f },
            { "height", 250.0f },
            { "rotation", 90.0f },
            { "name", "TestShape" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        var shape = pres.Slides[0].Shapes[0];
        Assert.Equal(50, shape.X);
        Assert.Equal(60, shape.Y);
        Assert.Equal(400, shape.Width);
        Assert.Equal(250, shape.Height);
        Assert.Equal(90, shape.Rotation);
        Assert.Equal("TestShape", shape.Name);
    }

    #endregion

    #region Position Updates

    [Fact]
    public void Execute_WithX_UpdatesXPosition()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "x", 300.0f }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(300, pres.Slides[0].Shapes[0].X);
    }

    [Fact]
    public void Execute_WithY_UpdatesYPosition()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "y", 400.0f }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(400, pres.Slides[0].Shapes[0].Y);
    }

    #endregion

    #region Size Updates

    [Fact]
    public void Execute_WithWidth_UpdatesWidth()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "width", 500.0f }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(500, pres.Slides[0].Shapes[0].Width);
    }

    [Fact]
    public void Execute_WithHeight_UpdatesHeight()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "height", 300.0f }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(300, pres.Slides[0].Shapes[0].Height);
    }

    #endregion

    #region AlternativeText

    [Fact]
    public void Execute_WithAlternativeText_UpdatesAlternativeText()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "alternativeText", "Description of shape" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("Description of shape", pres.Slides[0].Shapes[0].AlternativeText);
    }

    [Fact]
    public void Execute_WithEmptyAlternativeText_ClearsAlternativeText()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.AlternativeText = "Old text";
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "alternativeText", "" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal("", pres.Slides[0].Shapes[0].AlternativeText);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
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
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
