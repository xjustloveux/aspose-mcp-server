using System.Drawing;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class SetPptShapeFormatHandlerTests : PptHandlerTestBase
{
    private readonly SetPptShapeFormatHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_SetFormat()
    {
        Assert.Equal("set_format", _handler.Operation);
    }

    #endregion

    #region Basic Format Operations

    [Fact]
    public void Execute_SetsFormat()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fillColor", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Format", result);
        AssertModified(context);
    }

    #endregion

    #region Line Color

    [Fact]
    public void Execute_WithLineColor_SetsLineFormat()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "lineColor", "#FF0000" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.Solid, shape.LineFormat.FillFormat.FillType);
        AssertModified(context);
    }

    #endregion

    #region Transparency

    [Fact]
    public void Execute_WithTransparency_SetsTransparency()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "transparency", 0.5f }
        });

        _handler.Execute(context, parameters);

        AssertModified(context);
    }

    #endregion

    #region Multiple Properties

    [Fact]
    public void Execute_WithMultipleProperties_AppliesAll()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fillColor", "#FF0000" },
            { "lineColor", "#00FF00" },
            { "lineWidth", 3.0f }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
        Assert.Equal(FillType.Solid, shape.LineFormat.FillFormat.FillType);
        Assert.Equal(3.0, shape.LineFormat.Width);
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
            { "shapeIndex", 0 },
            { "fillColor", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("shape 0", result);
        Assert.Contains("slide 0", result);
    }

    #endregion

    #region Fill Color

    [Fact]
    public void Execute_WithFillColor_SetsSolidFill()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fillColor", "#0000FF" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
    }

    [Theory]
    [InlineData("#FF0000")]
    [InlineData("#00FF00")]
    [InlineData("#0000FF")]
    [InlineData("Red")]
    [InlineData("Blue")]
    public void Execute_WithVariousFillColors_AppliesColor(string color)
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fillColor", color }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Format", result);
        AssertModified(context);
    }

    #endregion

    #region Line Width

    [Fact]
    public void Execute_WithLineWidth_SetsLineWidth()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "lineWidth", 5.0f }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(5.0, shape.LineFormat.Width);
        AssertModified(context);
    }

    [Theory]
    [InlineData(1.0f)]
    [InlineData(3.0f)]
    [InlineData(10.0f)]
    public void Execute_WithVariousLineWidths_AppliesWidth(float width)
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "lineWidth", width }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(width, shape.LineFormat.Width);
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_FormatsOnSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 },
            { "fillColor", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("1", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_DefaultSlideIndex_FormatsOnFirstSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fillColor", "#FF0000" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("0", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fillColor", "#FF0000" }
        });

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
            { "shapeIndex", invalidIndex },
            { "fillColor", "#FF0000" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
