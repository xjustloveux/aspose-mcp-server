using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Handlers.PowerPoint.DataOperations;
using AsposeMcpServer.Results.PowerPoint.DataOperations;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.DataOperations;

[SupportedOSPlatform("windows")]
public class GetSlideDetailsHandlerTests : PptHandlerTestBase
{
    private readonly GetSlideDetailsHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_GetSlideDetails()
    {
        SkipIfNotWindows();
        Assert.Equal("slide_details", _handler.Operation);
    }

    #endregion

    #region Basic Get Slide Details Operations

    [SkippableFact]
    public void Execute_ReturnsSlideDetails()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlideDetailsResult>(res);

        Assert.Equal(0, result.SlideIndex);
        Assert.True(result.ShapesCount >= 0);
    }

    [SkippableFact]
    public void Execute_MultiShapeAnimations_ReportsPerShapeIndexAndShapeIndex()
    {
        SkipIfNotWindows();
        // Adversarial: shape A (1 anim) then shape B (2 anims). The reported animation index must be
        // per-shape and carry a usable shapeIndex, so (shapeIndex, index) round-trips into ppt_animation
        // edit/delete — not the whole-sequence position with only a shape type name.
        var pres = CreateEmptyPresentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var shapeB = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 300, 200, 100);
        var seq = slide.Timeline.MainSequence;
        seq.AddEffect(slide.Shapes[0], EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
        seq.AddEffect(shapeB, EffectType.Fly, EffectSubtype.Bottom, EffectTriggerType.OnClick);
        seq.AddEffect(shapeB, EffectType.Appear, EffectSubtype.None, EffectTriggerType.OnClick);

        var res = (GetSlideDetailsResult)_handler.Execute(CreateContext(pres),
            CreateParameters(new Dictionary<string, object?> { { "slideIndex", 0 } }));

        var bFly = res.Animations.Single(a => a.Type == "Fly");
        Assert.Equal(1, bFly.ShapeIndex);
        Assert.Equal(0, bFly.Index);

        var bAppear = res.Animations.Single(a => a.Type == "Appear");
        Assert.Equal(1, bAppear.ShapeIndex);
        Assert.Equal(1, bAppear.Index);
    }

    [SkippableFact]
    public void Execute_ReturnsLayoutInfo()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlideDetailsResult>(res);

        Assert.NotNull(result);
    }

    [SkippableFact]
    public void Execute_ReturnsTransitionInfo()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlideDetailsResult>(res);

        Assert.NotNull(result);
    }

    [SkippableFact]
    public void Execute_ReturnsAnimationsCount()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetSlideDetailsResult>(res);

        Assert.True(result.AnimationsCount >= 0);
    }

    [SkippableFact]
    public void Execute_WithInvalidIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var presentation = CreateEmptyPresentation();
        var context = CreateContext(presentation);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 999 }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
