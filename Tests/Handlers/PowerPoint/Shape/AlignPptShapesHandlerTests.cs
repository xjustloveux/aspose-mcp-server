using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

[SupportedOSPlatform("windows")]
public class AlignPptShapesHandlerTests : PptHandlerTestBase
{
    private static readonly int[] TwoShapeIndices = [0, 1];
    private static readonly int[] ThreeShapeIndices = [0, 1, 2];
    private static readonly int[] SingleShapeIndex = [0];
    private static readonly int[] InvalidShapeIndices = [0, 10];

    private readonly AlignPptShapesHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Align()
    {
        SkipIfNotWindows();
        Assert.Equal("align", _handler.Operation);
    }

    #endregion

    #region Align Left

    [SkippableFact]
    public void Execute_AlignLeft_AlignsToLeftmost()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 50, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", TwoShapeIndices },
            { "align", "left" }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(50, pres.Slides[0].Shapes[0].X);
        Assert.Equal(50, pres.Slides[0].Shapes[1].X);
    }

    #endregion

    #region Align to Slide

    [SkippableFact]
    public void Execute_WithAlignToSlideTrue_AlignsToSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", TwoShapeIndices },
            { "align", "left" },
            { "alignToSlide", true }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(0, pres.Slides[0].Shapes[0].X);
        Assert.Equal(0, pres.Slides[0].Shapes[1].X);
    }

    #endregion

    #region Multiple Shapes

    [SkippableFact]
    public void Execute_WithThreeShapes_AlignsAll()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 100, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Triangle, 300, 300, 100, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", ThreeShapeIndices },
            { "align", "top" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(100, pres.Slides[0].Shapes[0].Y);
        Assert.Equal(100, pres.Slides[0].Shapes[1].Y);
        Assert.Equal(100, pres.Slides[0].Shapes[2].Y);
    }

    #endregion

    #region Result Message

    [SkippableFact]
    public void Execute_AlignCenter_CentersShapesHorizontally()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", TwoShapeIndices },
            { "align", "center" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(pres.Slides[0].Shapes[0].X, pres.Slides[0].Shapes[1].X);
    }

    #endregion

    #region Basic Align Operations

    [SkippableFact]
    public void Execute_AlignsShapes()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", TwoShapeIndices },
            { "align", "left" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(100, pres.Slides[0].Shapes[0].X);
        Assert.Equal(100, pres.Slides[0].Shapes[1].X);
    }

    [SkippableTheory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("top")]
    [InlineData("middle")]
    [InlineData("bottom")]
    public void Execute_SupportsVariousAlignments(string align)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", TwoShapeIndices },
            { "align", align }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        var shape0 = pres.Slides[0].Shapes[0];
        var shape1 = pres.Slides[0].Shapes[1];
        switch (align)
        {
            case "left":
            case "center":
            case "right":
                Assert.Equal(shape0.X, shape1.X);
                break;
            case "top":
            case "middle":
            case "bottom":
                Assert.Equal(shape0.Y, shape1.Y);
                break;
        }
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndices", TwoShapeIndices },
            { "align", "left" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_WithoutShapeIndices_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithoutAlign_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", TwoShapeIndices }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("align", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_WithLessThanTwoShapes_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", SingleShapeIndex },
            { "align", "left" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("2", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidAlign_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", TwoShapeIndices },
            { "align", "invalid" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("align", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 200, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndices", InvalidShapeIndices },
            { "align", "left" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
