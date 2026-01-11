using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class GetPptShapesHandlerTests : PptHandlerTestBase
{
    private readonly GetPptShapesHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetShapes()
    {
        Assert.Equal("get_shapes", _handler.Operation);
    }

    #endregion

    #region Shape Properties

    [Fact]
    public void Execute_ReturnsShapeProperties()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 300, 400);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        var shapes = json.RootElement.GetProperty("shapes");
        Assert.True(shapes.GetArrayLength() > 0);
        var firstShape = shapes[0];
        Assert.True(firstShape.TryGetProperty("index", out _));
        Assert.True(firstShape.TryGetProperty("name", out _));
        Assert.True(firstShape.TryGetProperty("type", out _));
        Assert.True(firstShape.TryGetProperty("x", out _));
        Assert.True(firstShape.TryGetProperty("y", out _));
        Assert.True(firstShape.TryGetProperty("width", out _));
        Assert.True(firstShape.TryGetProperty("height", out _));
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyPresentation()
    {
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

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException(int invalidIndex)
    {
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

    [Fact]
    public void Execute_ReturnsShapesInfo()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("count", out _));
        Assert.True(json.RootElement.TryGetProperty("shapes", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectCount()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 300, 100, 200, 100);
        var shapeCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(shapeCount, json.RootElement.GetProperty("count").GetInt32());
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_GetsFromSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(1, json.RootElement.GetProperty("slideIndex").GetInt32());
    }

    [Fact]
    public void Execute_DefaultSlideIndex_GetsFromFirstSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
    }

    #endregion
}
