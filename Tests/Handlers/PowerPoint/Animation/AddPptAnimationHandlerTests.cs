using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Handlers.PowerPoint.Animation;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Animation;

[SupportedOSPlatform("windows")]
public class AddPptAnimationHandlerTests : PptHandlerTestBase
{
    private readonly AddPptAnimationHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Add()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_AddsAnimation()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var animations = pres.Slides[0].Timeline.MainSequence;
            Assert.True(animations.Count > 0, "Animation should be added");
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_DefaultsToFadeEffect()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var slide = pres.Slides[0];
            var animations = slide.Timeline.MainSequence;
            Assert.True(animations.Count > 0, "Animation should be added");
            Assert.Equal(EffectType.Fade, animations[0].Type);
        }

        AssertModified(context);
    }

    #endregion

    #region Custom Effect Types

    [SkippableTheory]
    [InlineData("Fly", EffectType.Fly)]
    [InlineData("Wipe", EffectType.Wipe)]
    [InlineData("Appear", EffectType.Appear)]
    [InlineData("Fade", EffectType.Fade)]
    public void Execute_WithEffectType_AddsCorrectEffect(string effectTypeStr, EffectType expectedType)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "effectType", effectTypeStr }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var animations = pres.Slides[0].Timeline.MainSequence;
            Assert.True(animations.Count > 0, "Animation should be added");
            Assert.Equal(expectedType, animations[0].Type);
        }

        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithEffectSubtype_AddsAnimation()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "effectType", "Fly" },
            { "effectSubtype", "FromBottom" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var animations = pres.Slides[0].Timeline.MainSequence;
            Assert.True(animations.Count > 0);
            Assert.Equal(EffectSubtype.Bottom, animations[0].Subtype);
        }
    }

    [SkippableFact]
    public void Execute_WithTriggerType_AddsAnimation()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "triggerType", "OnClick" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var animations = pres.Slides[0].Timeline.MainSequence;
            Assert.True(animations.Count > 0);
            Assert.Equal(EffectTriggerType.OnClick, animations[0].Timing.TriggerType);
        }
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "shapeIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("shapeIndex", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithNegativeShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithInvalidEffectType_DefaultsToFade()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "effectType", "InvalidEffect" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var animations = pres.Slides[0].Timeline.MainSequence;
            Assert.True(animations.Count > 0, "Animation should be added with default effect type");
        }
    }

    #endregion
}
