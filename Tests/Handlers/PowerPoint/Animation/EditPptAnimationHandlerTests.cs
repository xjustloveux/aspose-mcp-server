using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Handlers.PowerPoint.Animation;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Animation;

[SupportedOSPlatform("windows")]
public class EditPptAnimationHandlerTests : PptHandlerTestBase
{
    private readonly EditPptAnimationHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Edit()
    {
        SkipIfNotWindows();
        Assert.Equal("edit", _handler.Operation);
    }

    #endregion

    #region Replace All Animations

    [SkippableFact]
    public void Execute_WithoutAnimationIndex_ReplacesAllAnimations()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "effectType", "Fly" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(EffectType.Fly, pres.Slides[0].Timeline.MainSequence[0].Type);
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

    #region Clear Animations

    [SkippableFact]
    public void Execute_WithoutAnimationIndexAndWithoutEffect_ClearsAnimations()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(0, pres.Slides[0].Timeline.MainSequence.Count);
    }

    #endregion

    #region Basic Edit Operations

    [SkippableFact]
    public void Execute_EditsAnimation()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", 2.0f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(2.0f, pres.Slides[0].Timeline.MainSequence[0].Timing.Duration);
        AssertModified(context);
    }

    #endregion

    #region Edit Duration and Delay

    [SkippableTheory]
    [InlineData(1.0f)]
    [InlineData(2.5f)]
    [InlineData(5.0f)]
    public void Execute_WithDuration_ChangesDuration(float duration)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", duration }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(duration, pres.Slides[0].Timeline.MainSequence[0].Timing.Duration);
    }

    [SkippableFact]
    public void Execute_WithDelay_ChangesDelay()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "delay", 1.5f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(1.5f, pres.Slides[0].Timeline.MainSequence[0].Timing.TriggerDelayTime);
    }

    #endregion

    #region Boundary Condition Tests

    [SkippableTheory]
    [InlineData(0.0f)]
    [InlineData(0.001f)]
    [InlineData(100.0f)]
    public void Execute_WithBoundaryDuration_AcceptsValidValues(float duration)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", duration }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(duration, pres.Slides[0].Timeline.MainSequence[0].Timing.Duration);
    }

    [SkippableTheory]
    [InlineData(-1.0f)]
    [InlineData(-0.5f)]
    public void Execute_WithNegativeDuration_StillUpdatesAnimation(float duration)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "duration", duration }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [SkippableTheory]
    [InlineData(0.0f)]
    [InlineData(0.001f)]
    [InlineData(100.0f)]
    public void Execute_WithBoundaryDelay_AcceptsValidValues(float delay)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "delay", delay }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(delay, pres.Slides[0].Timeline.MainSequence[0].Timing.TriggerDelayTime);
    }

    [SkippableTheory]
    [InlineData(-1.0f)]
    [InlineData(-0.5f)]
    public void Execute_WithNegativeDelay_StillUpdatesAnimation(float delay)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "delay", delay }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithNegativeShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithNegativeAnimationIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
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
        var pres = CreatePresentationWithAnimation();
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

    [SkippableFact]
    public void Execute_WithInvalidShapeIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    [SkippableFact]
    public void Execute_WithInvalidAnimationIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
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

    #region Effect Type Changes

    [SkippableFact]
    public void Execute_WithEffectType_ChangesEffect()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "effectType", "Fly" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode()) Assert.Equal(EffectType.Fly, pres.Slides[0].Timeline.MainSequence[0].Type);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithEffectSubtype_ChangesSubtype()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "effectType", "Fly" },
            { "effectSubtype", "FromLeft" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithTriggerType_ChangesTrigger()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "triggerType", "WithPrevious" }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
            Assert.Equal(EffectTriggerType.WithPrevious, pres.Slides[0].Timeline.MainSequence[0].Timing.TriggerType);
    }

    [SkippableFact]
    public void Execute_WithAllEffectParameters_ChangesAll()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 },
            { "effectType", "Zoom" },
            { "effectSubtype", "In" },
            { "triggerType", "AfterPrevious" },
            { "duration", 1.5f },
            { "delay", 0.5f }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        if (!IsEvaluationMode())
        {
            var effect = pres.Slides[0].Timeline.MainSequence[0];
            Assert.Equal(1.5f, effect.Timing.Duration);
            Assert.Equal(0.5f, effect.Timing.TriggerDelayTime);
        }

        AssertModified(context);
    }

    #endregion
}
