using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Handlers.PowerPoint.Animation;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Animation;

public class AddPptAnimationHandlerTests : PptHandlerTestBase
{
    private readonly AddPptAnimationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Add()
    {
        Assert.Equal("add", _handler.Operation);
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

    #endregion

    #region Basic Add Operations

    [Fact]
    public void Execute_AddsAnimation()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("Animation", result);
        Assert.Contains("added", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide 0", result);
    }

    [Fact]
    public void Execute_ReturnsShapeIndex()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("shape 0", result);
    }

    [Fact]
    public void Execute_DefaultsToFadeEffect()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        _handler.Execute(context, parameters);

        var slide = pres.Slides[0];
        var animations = slide.Timeline.MainSequence;
        Assert.True(animations.Count > 0, "Animation should be added");
        Assert.Equal(EffectType.Fade, animations[0].Type);
        AssertModified(context);
    }

    #endregion

    #region Custom Effect Types

    [Theory]
    [InlineData("Fly", EffectType.Fly)]
    [InlineData("Wipe", EffectType.Wipe)]
    [InlineData("Appear", EffectType.Appear)]
    [InlineData("Fade", EffectType.Fade)]
    public void Execute_WithEffectType_AddsCorrectEffect(string effectTypeStr, EffectType expectedType)
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "effectType", effectTypeStr }
        });

        _handler.Execute(context, parameters);

        var animations = pres.Slides[0].Timeline.MainSequence;
        Assert.True(animations.Count > 0, "Animation should be added");
        Assert.Equal(expectedType, animations[0].Type);
        AssertModified(context);
    }

    [Fact]
    public void Execute_WithEffectSubtype_AddsAnimation()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "effectType", "Fly" },
            { "effectSubtype", "FromBottom" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result);
    }

    [Fact]
    public void Execute_WithTriggerType_AddsAnimation()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "triggerType", "OnClick" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Slide index", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Shape index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", -1 },
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Slide index", ex.Message);
    }

    [Fact]
    public void Execute_WithNegativeShapeIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Shape index", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidEffectType_DefaultsToFade()
    {
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "effectType", "InvalidEffect" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("added", result);
        var animations = pres.Slides[0].Timeline.MainSequence;
        Assert.True(animations.Count > 0, "Animation should be added with default effect type");
    }

    #endregion
}
