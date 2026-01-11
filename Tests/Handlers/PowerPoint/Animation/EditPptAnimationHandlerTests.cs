using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Handlers.PowerPoint.Animation;
using AsposeMcpServer.Tests.Helpers;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Animation;

public class EditPptAnimationHandlerTests : PptHandlerTestBase
{
    private readonly EditPptAnimationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Edit()
    {
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Replace All Animations

    [Fact]
    public void Execute_WithoutAnimationIndex_ReplacesAllAnimations()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "effectType", "Fly" }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    #endregion

    #region Helper Methods

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

    #region Basic Edit Operations

    [Fact]
    public void Execute_EditsAnimation()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", 2.0f }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
        AssertModified(context);
    }

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", 2.0f }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("slide 0", result);
    }

    [Fact]
    public void Execute_ReturnsShapeIndex()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", 2.0f }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("shape 0", result);
    }

    #endregion

    #region Edit Duration and Delay

    [Theory]
    [InlineData(1.0f)]
    [InlineData(2.5f)]
    [InlineData(5.0f)]
    public void Execute_WithDuration_ChangesDuration(float duration)
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", duration }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    [Fact]
    public void Execute_WithDelay_ChangesDelay()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "delay", 1.5f }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    #endregion

    #region Boundary Condition Tests

    [Theory]
    [InlineData(0.0f)]
    [InlineData(0.001f)]
    [InlineData(100.0f)]
    public void Execute_WithBoundaryDuration_AcceptsValidValues(float duration)
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", duration }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    [Theory]
    [InlineData(-1.0f)]
    [InlineData(-0.5f)]
    public void Execute_WithNegativeDuration_StillUpdatesAnimation(float duration)
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", duration }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    [Theory]
    [InlineData(0.0f)]
    [InlineData(0.001f)]
    [InlineData(100.0f)]
    public void Execute_WithBoundaryDelay_AcceptsValidValues(float delay)
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "delay", delay }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    [Theory]
    [InlineData(-1.0f)]
    [InlineData(-0.5f)]
    public void Execute_WithNegativeDelay_StillUpdatesAnimation(float delay)
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "delay", delay }
        });

        var result = _handler.Execute(context, parameters);

        Assert.Contains("updated", result);
    }

    [Fact]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithAnimation();
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
        var pres = CreatePresentationWithAnimation();
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
    public void Execute_WithNegativeAnimationIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("animationIndex", ex.Message);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithAnimation();
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
        var pres = CreatePresentationWithAnimation();
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
        var pres = CreatePresentationWithAnimation();
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
        var pres = CreatePresentationWithAnimation();
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
    public void Execute_WithInvalidAnimationIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("animationIndex", ex.Message);
    }

    #endregion
}
