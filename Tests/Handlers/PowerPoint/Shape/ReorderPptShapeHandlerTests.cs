using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

[SupportedOSPlatform("windows")]
public class ReorderPptShapeHandlerTests : PptHandlerTestBase
{
    private readonly ReorderPptShapeHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Reorder()
    {
        SkipIfNotWindows();
        Assert.Equal("reorder", _handler.Operation);
    }

    #endregion

    #region Preserve Shape Count

    [SkippableFact]
    public void Execute_PreservesShapeCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var initialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "toIndex", 1 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, pres.Slides[0].Shapes.Count);
    }

    #endregion

    #region Slide Index

    [SkippableFact]
    public void Execute_WithSlideIndex_ReordersOnSpecificSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var firstShape = pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var firstShapeName = firstShape.Name;
        var fromIndex = pres.Slides[1].Shapes.IndexOf(firstShape);
        var toIndex = fromIndex + 1;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", fromIndex },
            { "toIndex", toIndex }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(firstShapeName, pres.Slides[1].Shapes[toIndex].Name);
        AssertModified(context);
    }

    #endregion

    #region Result Message

    [SkippableFact]
    public void Execute_ReturnsIndicesInMessage()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var firstShape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var firstShapeName = firstShape.Name;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "toIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(firstShapeName, pres.Slides[0].Shapes[1].Name);
    }

    #endregion

    #region Basic Reorder Operations

    [SkippableFact]
    public void Execute_ReordersShape()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        var firstShape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Triangle, 300, 100, 200, 100);
        var firstShapeName = firstShape.Name;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "toIndex", 2 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(firstShapeName, pres.Slides[0].Shapes[2].Name);
        Assert.Equal(3, pres.Slides[0].Shapes.Count);
        AssertModified(context);
    }

    [SkippableTheory]
    [InlineData(0, 1)]
    [InlineData(0, 2)]
    [InlineData(2, 0)]
    [InlineData(1, 0)]
    public void Execute_ReordersToVariousPositions(int from, int to)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Triangle, 300, 100, 200, 100);
        var movedShapeName = pres.Slides[0].Shapes[from].Name;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", from },
            { "toIndex", to }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.Equal(movedShapeName, pres.Slides[0].Shapes[to].Name);
        Assert.Equal(3, pres.Slides[0].Shapes.Count);
        AssertModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 },
            { "toIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "toIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_WithoutToIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("toIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableTheory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", invalidIndex },
            { "toIndex", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableTheory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidToIndex_ThrowsArgumentException(int invalidIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 200, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "toIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
