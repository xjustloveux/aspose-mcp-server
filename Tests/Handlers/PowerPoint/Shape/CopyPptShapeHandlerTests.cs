using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

[SupportedOSPlatform("windows")]
public class CopyPptShapeHandlerTests : PptHandlerTestBase
{
    private readonly CopyPptShapeHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Copy()
    {
        SkipIfNotWindows();
        Assert.Equal("copy", _handler.Operation);
    }

    #endregion

    #region Preserve Source

    [SkippableFact]
    public void Execute_PreservesSourceShape()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var sourceInitialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromSlide", 0 },
            { "toSlide", 1 },
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(sourceInitialCount, pres.Slides[0].Shapes.Count);
    }

    #endregion

    #region Result Message

    [SkippableFact]
    public void Execute_CopyAddsShapeToTargetSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var targetInitialCount = pres.Slides[1].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromSlide", 0 },
            { "toSlide", 1 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(targetInitialCount + 1, pres.Slides[1].Shapes.Count);
    }

    #endregion

    #region Basic Copy Operations

    [SkippableFact]
    public void Execute_CopiesShapeToAnotherSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var targetInitialCount = pres.Slides[1].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromSlide", 0 },
            { "toSlide", 1 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
        Assert.Equal(targetInitialCount + 1, pres.Slides[1].Shapes.Count);
    }

    [SkippableTheory]
    [InlineData(0, 1)]
    [InlineData(0, 2)]
    [InlineData(1, 0)]
    public void Execute_CopiesBetweenVariousSlides(int from, int to)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[from].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var targetInitialCount = pres.Slides[to].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromSlide", from },
            { "toSlide", to },
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(targetInitialCount + 1, pres.Slides[to].Shapes.Count);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutFromSlide_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "toSlide", 1 },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("fromSlide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_WithoutToSlide_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromSlide", 0 },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("toSlide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromSlide", 0 },
            { "toSlide", 1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableTheory]
    [InlineData(-1, 1)]
    [InlineData(10, 1)]
    public void Execute_WithInvalidFromSlide_ThrowsArgumentException(int invalidFrom, int to)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromSlide", invalidFrom },
            { "toSlide", to },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("fromSlide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableTheory]
    [InlineData(0, -1)]
    [InlineData(0, 10)]
    public void Execute_WithInvalidToSlide_ThrowsArgumentException(int from, int invalidTo)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromSlide", from },
            { "toSlide", invalidTo },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("toSlide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableTheory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException(int invalidIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(2);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "fromSlide", 0 },
            { "toSlide", 1 },
            { "shapeIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
