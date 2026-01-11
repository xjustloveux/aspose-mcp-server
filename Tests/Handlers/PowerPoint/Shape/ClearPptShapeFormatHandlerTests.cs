using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class ClearPptShapeFormatHandlerTests : PptHandlerTestBase
{
    private readonly ClearPptShapeFormatHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_ClearFormat()
    {
        Assert.Equal("clear_format", _handler.Operation);
    }

    #endregion

    #region Basic Clear Operations

    [Fact]
    public void Execute_ClearsFormat()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("cleared", result, StringComparison.OrdinalIgnoreCase);
        AssertModified(context);
    }

    #endregion

    #region Result Message

    [Fact]
    public void Execute_ReturnsSlideAndShapeIndexInMessage()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("shape 0", result);
        Assert.Contains("slide 0", result);
    }

    #endregion

    #region Clear Fill

    [Fact]
    public void Execute_DefaultClearsFill()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.NoFill, shape.FillFormat.FillType);
    }

    [Fact]
    public void Execute_WithClearFillFalse_PreservesFill()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "clearFill", false }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
    }

    #endregion

    #region Clear Line

    [Fact]
    public void Execute_DefaultClearsLine()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.LineFormat.FillFormat.FillType = FillType.Solid;
        shape.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.NoFill, shape.LineFormat.FillFormat.FillType);
    }

    [Fact]
    public void Execute_WithClearLineFalse_PreservesLine()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.LineFormat.FillFormat.FillType = FillType.Solid;
        shape.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "clearLine", false }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.Solid, shape.LineFormat.FillFormat.FillType);
    }

    #endregion

    #region Clear Both

    [Fact]
    public void Execute_ClearsBothFillAndLine()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        shape.LineFormat.FillFormat.FillType = FillType.Solid;
        shape.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "clearFill", true },
            { "clearLine", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.NoFill, shape.FillFormat.FillType);
        Assert.Equal(FillType.NoFill, shape.LineFormat.FillFormat.FillType);
    }

    [Fact]
    public void Execute_WithBothFalse_PreservesBoth()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        shape.LineFormat.FillFormat.FillType = FillType.Solid;
        shape.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "clearFill", false },
            { "clearLine", false }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
        Assert.Equal(FillType.Solid, shape.LineFormat.FillFormat.FillType);
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_ClearsOnSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var shape = pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1", result);
        Assert.Equal(FillType.NoFill, pres.Slides[1].Shapes[0].FillFormat.FillType);
    }

    [Fact]
    public void Execute_DefaultSlideIndex_ClearsOnFirstSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0", result);
        Assert.Equal(FillType.NoFill, pres.Slides[0].Shapes[0].FillFormat.FillType);
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
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
