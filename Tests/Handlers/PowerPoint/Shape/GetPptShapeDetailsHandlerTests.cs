using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Shape;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Shape;

public class GetPptShapeDetailsHandlerTests : PptHandlerTestBase
{
    private readonly GetPptShapeDetailsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_GetShapeDetails()
    {
        Assert.Equal("get_shape_details", _handler.Operation);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_DoesNotModifyPresentation()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var initialCount = pres.Slides[0].Shapes.Count;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        Assert.Equal(initialCount, pres.Slides[0].Shapes.Count);
        AssertNotModified(context);
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_ReturnsShapeDetails()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 300, 400);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("index", out _));
        Assert.True(json.RootElement.TryGetProperty("name", out _));
        Assert.True(json.RootElement.TryGetProperty("type", out _));
        Assert.True(json.RootElement.TryGetProperty("x", out _));
        Assert.True(json.RootElement.TryGetProperty("y", out _));
        Assert.True(json.RootElement.TryGetProperty("width", out _));
        Assert.True(json.RootElement.TryGetProperty("height", out _));
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_ReturnsCorrectPosition()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 150, 250, 300, 400);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(150, json.RootElement.GetProperty("x").GetSingle());
        Assert.Equal(250, json.RootElement.GetProperty("y").GetSingle());
    }

    [Fact]
    public void Execute_ReturnsCorrectDimensions()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 350, 450);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(350, json.RootElement.GetProperty("width").GetSingle());
        Assert.Equal(450, json.RootElement.GetProperty("height").GetSingle());
    }

    #endregion

    #region AutoShape Details

    [Fact]
    public void Execute_ForAutoShape_ReturnsShapeType()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("shapeType", out var shapeType));
        Assert.Equal("Ellipse", shapeType.GetString());
    }

    [Fact]
    public void Execute_ForAutoShape_ReturnsFillType()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.FillFormat.FillType = FillType.Solid;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("fillType", out var fillType));
        Assert.Equal("Solid", fillType.GetString());
    }

    #endregion

    #region Rotation and Hidden

    [Fact]
    public void Execute_ReturnsRotation()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.Rotation = 45;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(45, json.RootElement.GetProperty("rotation").GetSingle());
    }

    [Fact]
    public void Execute_ReturnsHiddenStatus()
    {
        var pres = CreatePresentationWithSlides(1);
        var shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        shape.Hidden = true;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("hidden").GetBoolean());
    }

    #endregion

    #region Slide Index

    [Fact]
    public void Execute_WithSlideIndex_GetsFromSpecificSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[1].Shapes.AddAutoShape(ShapeType.Triangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 1 },
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("index").GetInt32());
    }

    [Fact]
    public void Execute_DefaultSlideIndex_GetsFromFirstSlide()
    {
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        var json = JsonDocument.Parse(result);
        Assert.Equal(0, json.RootElement.GetProperty("index").GetInt32());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithSlides(1);
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
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
        pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 200);
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
