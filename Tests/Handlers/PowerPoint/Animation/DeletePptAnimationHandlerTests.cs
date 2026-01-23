using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Handlers.PowerPoint.Animation;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Animation;

public class DeletePptAnimationHandlerTests : PptHandlerTestBase
{
    private readonly DeletePptAnimationHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Delete()
    {
        Assert.Equal("delete", _handler.Operation);
    }

    #endregion

    #region Delete All From Shape

    [Fact]
    public void Execute_WithoutAnimationIndex_DeletesAllFromShape()
    {
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

    [Fact]
    public void Execute_WithoutShapeIndex_ClearsAllAnimations()
    {
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

    [Fact]
    public void Execute_DeletesAnimation()
    {
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

    [Fact]
    public void Execute_ReturnsSlideIndex()
    {
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

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 }
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

    [Fact]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", -1 }
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

    [Theory]
    [InlineData(int.MaxValue)]
    public void Execute_WithExtremeSlideIndex_ThrowsArgumentException(int slideIndex)
    {
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
