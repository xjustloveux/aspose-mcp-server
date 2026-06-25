using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Handlers.PowerPoint.Animation;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Animation;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Animation;

[SupportedOSPlatform("windows")]
public class DeletePptAnimationHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptAnimationHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Delete()
    {
        SkipIfNotWindows();
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Delete All From Shape

    [SkippableFact]
    public void Execute_WithoutAnimationIndex_DeletesAllFromShape()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithMultipleAnimations();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
    }

    #endregion

    #region Delete All From Slide

    [SkippableFact]
    public void Execute_WithoutShapeIndex_ClearsAllAnimations()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
        Assert.Equal(0, pres.Slides[0].Timeline.MainSequence.Count);
    }

    #endregion

    #region Basic Delete Operations

    [SkippableFact]
    public void Execute_DeletesAnimation()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", 0 },
            { "animationIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("deleted", result.Message);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_ReturnsSlideIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<SuccessResult>(res);

        Assert.Contains("slide 0", result.Message);
    }

    [SkippableFact]
    public void Execute_RoundTrip_MultiShape_GetReportedIndexDeletesThatAnimation()
    {
        SkipIfNotWindows();
        // Adversarial: shape A has 1 animation, shape B has 2 (Fly then Fade). 'get' must report each
        // animation with an index that, paired with its shapeIndex, deletes THAT animation via delete.
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var shapeB = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 300, 200, 100);
        var seq = slide.Timeline.MainSequence;
        seq.AddEffect(slide.Shapes[0], EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
        seq.AddEffect(shapeB, EffectType.Fly, EffectSubtype.Bottom, EffectTriggerType.OnClick);
        seq.AddEffect(shapeB, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);

        var getRes = (GetAnimationsResult)new GetPptAnimationsHandler()
            .Execute(CreateContext(pres), CreateParameters(new Dictionary<string, object?> { { "slideIndex", 0 } }));
        var flyOnB = getRes.Animations.Single(a => a is { ShapeIndex: 1, EffectType: "Fly" });

        _handler.Execute(CreateContext(pres), CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", flyOnB.ShapeIndex },
            { "animationIndex", flyOnB.Index }
        }));

        var remainingOnB = slide.Timeline.MainSequence.Where(e => e.TargetShape == shapeB).ToList();
        Assert.Single(remainingOnB);
        Assert.Equal(EffectType.Fade, remainingOnB[0].Type);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 }
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

    [SkippableFact]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", -1 }
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

    [SkippableTheory]
    [InlineData(int.MaxValue)]
    public void Execute_WithExtremeSlideIndex_ThrowsArgumentException(int slideIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
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

    private static Presentation CreatePresentationWithAnimation()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        slide.Timeline.MainSequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None,
            EffectTriggerType.OnClick);
        return pres;
    }

    private static Presentation CreatePresentationWithMultipleAnimations()
    {
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        slide.Timeline.MainSequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None,
            EffectTriggerType.OnClick);
        slide.Timeline.MainSequence.AddEffect(shape, EffectType.Fly, EffectSubtype.Bottom,
            EffectTriggerType.AfterPrevious);
        return pres;
    }

    #endregion
}
