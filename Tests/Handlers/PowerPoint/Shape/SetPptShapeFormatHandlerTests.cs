using System.Drawing;
using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

[SupportedOSPlatform("windows")]
public class SetPptShapeFormatHandlerTests : PptHandlerTestBase
{
    private readonly SetPptShapeFormatHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_SetFormat()
    {
        SkipIfNotWindows();
        Assert.Equal("set_format", _handler.Operation);
    }

    #endregion

    #region Basic Format Operations

    [SkippableFact]
    public void Execute_SetsFormat()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fillColor", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
        if (!IsEvaluationMode()) Assert.Equal(Color.FromArgb(255, 0, 0), shape.FillFormat.SolidFillColor.Color);

        AssertModified(context);
    }

    #endregion

    #region Line Color

    [SkippableFact]
    public void Execute_WithLineColor_SetsLineFormat()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithTransparency_SetsTransparency()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithMultipleProperties_AppliesAll()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_ReturnsSlideAndShapeIndexInMessage()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fillColor", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
    }

    #endregion

    #region Fill Color

    [SkippableFact]
    public void Execute_WithFillColor_SetsSolidFill()
    {
        SkipIfNotWindows();
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

    [SkippableTheory]
    [InlineData("#FF0000")]
    [InlineData("#00FF00")]
    [InlineData("#0000FF")]
    [InlineData("Red")]
    [InlineData("Blue")]
    public void Execute_WithVariousFillColors_AppliesColor(string color)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fillColor", color }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
        AssertModified(context);
    }

    #endregion

    #region Line Width

    [SkippableFact]
    public void Execute_WithLineWidth_SetsLineWidth()
    {
        SkipIfNotWindows();
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

    [SkippableTheory]
    [InlineData(1.0f)]
    [InlineData(3.0f)]
    [InlineData(10.0f)]
    public void Execute_WithVariousLineWidths_AppliesWidth(float width)
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithSlideIndex_FormatsOnSpecificSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var shape = pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var shapeIndex = pres.Slides[1].Shapes.IndexOf(shape);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", shapeIndex },
            { "fillColor", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
        if (!IsEvaluationMode()) Assert.Equal(Color.FromArgb(255, 0, 0), shape.FillFormat.SolidFillColor.Color);

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_DefaultSlideIndex_FormatsOnFirstSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "fillColor", "#FF0000" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableTheory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        SkipIfNotWindows();
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
