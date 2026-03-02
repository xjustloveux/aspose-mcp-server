using System.Runtime.Versioning;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Results.PowerPoint.Shape;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

[SupportedOSPlatform("windows")]
public class GetPptShapesHandlerTests : PptHandlerTestBase
{
    private readonly GetPptShapesHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetShapes()
    {
        SkipIfNotWindows();
        Assert.Equal("get_shapes", _handler.Operation);
    }

    #endregion

    #region Shape Properties

    [SkippableFact]
    public void Execute_ReturnsShapeProperties()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 300, 400);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesResult>(res);

        Assert.True(result.Shapes.Count > 0);
        var firstShape = result.Shapes[0];
        Assert.True(firstShape.Index >= 0);
        Assert.NotNull(firstShape.Type);
        Assert.True(firstShape.X >= 0);
        Assert.True(firstShape.Y >= 0);
        Assert.True(firstShape.Width >= 0);
        Assert.True(firstShape.Height >= 0);
    }

    #endregion

    #region Read-Only Verification

    [SkippableFact]
    public void Execute_DoesNotModifyPresentation()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var initialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, pres.Slides[0].Shapes.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Error Handling

    [SkippableTheory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException(int invalidIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", invalidIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("index", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Basic Get Operations

    [SkippableFact]
    public void Execute_ReturnsShapesInfo()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesResult>(res);

        Assert.Equal(0, result.SlideIndex);
        Assert.True(result.Count > 0);
        Assert.NotNull(result.Shapes);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsCorrectCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 300, 100, 200, 100);
        var shapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesResult>(res);

        Assert.Equal(shapeCount, result.Count);
    }

    #endregion

    #region Slide Index

    [SkippableFact]
    public void Execute_WithSlideIndex_GetsFromSpecificSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesResult>(res);

        Assert.Equal(1, result.SlideIndex);
    }

    [SkippableFact]
    public void Execute_DefaultSlideIndex_GetsFromFirstSlide()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetShapesResult>(res);

        Assert.Equal(0, result.SlideIndex);
    }

    #endregion
}
