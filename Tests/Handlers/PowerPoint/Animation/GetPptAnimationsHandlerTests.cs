using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Handlers.PowerPoint.Animation;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Animation;

public class GetPptAnimationsHandlerTests : PptHandlerTestBase
{
    private readonly GetPptAnimationsHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Get()
    {
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Filter By Shape

    [Fact]
    public void Execute_WithShapeIndex_FiltersAnimations()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("filterByShapeIndex", out var filter));
        Assert.Equal(0, filter.GetInt32());
    }

    #endregion

    #region No Animations

    [Fact]
    public void Execute_NoAnimations_ReturnsEmptyArray()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("totalAnimationsOnSlide").GetInt32());
        Assert.Equal(0, json.RootElement.GetProperty("animations").GetArrayLength());
    }

    #endregion

    #region Basic Get Operations

    [Fact]
    public void Execute_GetsAnimations()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(0, json.RootElement.GetProperty("slideIndex").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsTotalCount()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.Equal(1, json.RootElement.GetProperty("totalAnimationsOnSlide").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsAnimationsArray()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("animations", out var animations));
        Assert.Equal(JsonValueKind.Array, animations.ValueKind);
    }

    #endregion

    #region Animation Details

    [Fact]
    public void Execute_ReturnsAnimationIndex()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstAnimation = json.RootElement.GetProperty("animations")[0];

        Assert.Equal(0, firstAnimation.GetProperty("index").GetInt32());
    }

    [Fact]
    public void Execute_ReturnsShapeIndex()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstAnimation = json.RootElement.GetProperty("animations")[0];

        Assert.True(firstAnimation.TryGetProperty("shapeIndex", out _));
    }

    [Fact]
    public void Execute_ReturnsEffectType()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);
        var firstAnimation = json.RootElement.GetProperty("animations")[0];

        Assert.NotNull(firstAnimation.GetProperty("effectType").GetString());
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Slide index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Slide index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeShapeIndex_ReturnsAnimationsWithFilter()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", -1 }
        });

        var result = _handler.Execute(context, parameters);
        var json = JsonDocument.Parse(result);

        Assert.True(json.RootElement.TryGetProperty("filterByShapeIndex", out var filter));
        Assert.Equal(-1, filter.GetInt32());
    }

    [Theory]
    [InlineData(int.MaxValue)]
    public void Execute_WithExtremeSlideIndex_ThrowsArgumentException(int slideIndex)
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", slideIndex }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Slide index", ex.Message);
    }

    #endregion

    #region Helper Methods

    private static Presentation CreatePresentationWithShape()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        return pres;
    }

    private static Presentation CreatePresentationWithAnimation()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        slide.Timeline.MainSequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None,
            EffectTriggerType.OnClick);
        return pres;
    }

    #endregion
}
