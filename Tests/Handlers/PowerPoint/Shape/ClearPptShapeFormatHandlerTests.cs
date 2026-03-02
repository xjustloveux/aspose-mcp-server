using System.Drawing;
using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

[SupportedOSPlatform("windows")]
public class ClearPptShapeFormatHandlerTests : PptHandlerTestBase
{
    private readonly ClearPptShapeFormatHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_ClearFormat()
    {
        SkipIfNotWindows();
        Assert.Equal("clear_format", _handler.Operation);
    }

    #endregion

    #region Basic Clear Operations

    [SkippableFact]
    public void Execute_ClearsFormat()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(FillType.NoFill, shape.FillFormat.FillType);
    }

    #endregion

    #region Result Message

    [SkippableFact]
    public void Execute_ClearsFillAndLineOnFirstSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        shape.LineFormat.FillFormat.FillType = FillType.Solid;
        shape.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(FillType.NoFill, shape.FillFormat.FillType);
        Assert.Equal(FillType.NoFill, shape.LineFormat.FillFormat.FillType);
    }

    #endregion

    #region Clear Fill

    [SkippableFact]
    public void Execute_DefaultClearsFill()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithClearFillFalse_PreservesFill()
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
            { "clearFill", false }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(FillType.Solid, shape.FillFormat.FillType);
    }

    #endregion

    #region Clear Line

    [SkippableFact]
    public void Execute_DefaultClearsLine()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithClearLineFalse_PreservesLine()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_ClearsBothFillAndLine()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithBothFalse_PreservesBoth()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithSlideIndex_ClearsOnSpecificSlide()
    {
        SkipIfNotWindows();
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

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(FillType.NoFill, pres.Slides[1].Shapes[0].FillFormat.FillType);
    }

    [SkippableFact]
    public void Execute_DefaultSlideIndex_ClearsOnFirstSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        shape.FillFormat.FillType = FillType.Solid;
        shape.FillFormat.SolidFillColor.Color = Color.Red;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(FillType.NoFill, pres.Slides[0].Shapes[0].FillFormat.FillType);
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
        var parameters = CreateEmptyParameters();

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
            { "shapeIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
