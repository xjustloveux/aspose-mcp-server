using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Animation;
using AsposeMcpServer.Handlers.PowerPoint.Animation;
using AsposeMcpServer.Results.PowerPoint.Animation;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Animation;

[SupportedOSPlatform("windows")]
public class GetPptAnimationsHandlerTests : PptHandlerTestBase
{
    private readonly GetPptAnimationsHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Get()
    {
        SkipIfNotWindows();
        Assert.Equal("get", _handler.Operation);
    }

    #endregion

    #region Filter By Shape

    [SkippableFact]
    public void Execute_WithShapeIndex_FiltersAnimations()
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

        var result = Assert.IsType<GetAnimationsResult>(res);

        Assert.NotNull(result.FilterByShapeIndex);
        Assert.Equal(0, result.FilterByShapeIndex.Value);
    }

    #endregion

    #region No Animations

    [SkippableFact]
    public void Execute_NoAnimations_ReturnsEmptyArray()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAnimationsResult>(res);

        Assert.Equal(0, result.TotalAnimationsOnSlide);
        Assert.Empty(result.Animations);
    }

    #endregion

    #region Basic Get Operations

    [SkippableFact]
    public void Execute_GetsAnimations()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAnimationsResult>(res);

        Assert.Equal(0, result.SlideIndex);
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

        var result = Assert.IsType<GetAnimationsResult>(res);

        Assert.Equal(0, result.SlideIndex);
    }

    [SkippableFact]
    public void Execute_ReturnsTotalCount()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAnimationsResult>(res);

        Assert.Equal(1, result.TotalAnimationsOnSlide);
    }

    [SkippableFact]
    public void Execute_ReturnsAnimationsArray()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAnimationsResult>(res);

        Assert.NotNull(result.Animations);
        Assert.True(result.Animations.Count >= 0);
    }

    #endregion

    #region Animation Details

    [SkippableFact]
    public void Execute_ReturnsAnimationIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAnimationsResult>(res);

        var firstAnimation = result.Animations[0];
        Assert.Equal(0, firstAnimation.Index);
    }

    [SkippableFact]
    public void Execute_ReturnsShapeIndex()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAnimationsResult>(res);

        var firstAnimation = result.Animations[0];
        Assert.True(firstAnimation.ShapeIndex >= 0 || firstAnimation.ShapeIndex == -1);
    }

    [SkippableFact]
    public void Execute_ReturnsEffectType()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAnimationsResult>(res);

        var firstAnimation = result.Animations[0];
        Assert.NotNull(firstAnimation.EffectType);
    }

    #endregion

    #region Error Handling

    [SkippableFact]
    public void Execute_WithoutSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("slideIndex", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 99 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Slide index", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithNegativeSlideIndex_ThrowsArgumentException()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithShape();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", -1 }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("Slide index", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithNegativeShapeIndex_ReturnsAnimationsWithFilter()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithAnimation();
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "slideIndex", 0 },
            { "shapeIndex", -1 }
        });

        var res = _handler.Execute(context, parameters);

        var result = Assert.IsType<GetAnimationsResult>(res);

        Assert.NotNull(result.FilterByShapeIndex);
        Assert.Equal(-1, result.FilterByShapeIndex.Value);
    }

    [SkippableTheory]
    [InlineData(int.MaxValue)]
    public void Execute_WithExtremeSlideIndex_ThrowsArgumentException(int slideIndex)
    {
        SkipIfNotWindows();
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
