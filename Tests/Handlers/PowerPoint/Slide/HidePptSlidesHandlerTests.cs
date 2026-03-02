using System.Runtime.Versioning;
using System.Text.Json;
using AsposeMcpServer.Handlers.PowerPoint.Slide;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Slide;

[SupportedOSPlatform("windows")]
public class HidePptSlidesHandlerTests : PptHandlerTestBase
{
    private static readonly int[] IndicesZeroTwo = [0, 2];
    private static readonly int[] IndicesZeroTwoFour = [0, 2, 4];
    private static readonly int[] IndicesOneThree = [1, 3];

    private readonly HidePptSlidesHandler _handler = new();

    #region Operation Property

    [SkippableFact]
    public void Operation_Returns_Hide()
    {
        SkipIfNotWindows();
        Assert.Equal("hide", _handler.Operation);
    }

    #endregion

    #region Single Slide

    [SkippableTheory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Execute_WithSingleIndex_HidesSingleSlide(int slideIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var indices = JsonSerializer.Serialize(new[] { slideIndex });
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", true },
            { "slideIndices", indices }
        });

        _handler.Execute(context, parameters);

        for (var i = 0; i < pres.Slides.Count; i++)
            Assert.Equal(i == slideIndex, pres.Slides[i].Hidden);
    }

    #endregion

    #region Result Message

    [SkippableFact]
    public void Execute_ReturnsSlideCountInMessage()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(5);
        var indices = JsonSerializer.Serialize(IndicesZeroTwo);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", true },
            { "slideIndices", indices }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Hidden);
        Assert.False(pres.Slides[1].Hidden);
        Assert.True(pres.Slides[2].Hidden);
        Assert.False(pres.Slides[3].Hidden);
        Assert.False(pres.Slides[4].Hidden);
        AssertModified(context);
    }

    #endregion

    #region Preserve Other Slides

    [SkippableFact]
    public void Execute_PreservesUnaffectedSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(5);
        pres.Slides[1].Hidden = true;
        pres.Slides[3].Hidden = true;
        var indices = JsonSerializer.Serialize(IndicesZeroTwo);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", true },
            { "slideIndices", indices }
        });

        _handler.Execute(context, parameters);

        Assert.True(pres.Slides[0].Hidden);
        Assert.True(pres.Slides[1].Hidden);
        Assert.True(pres.Slides[2].Hidden);
        Assert.True(pres.Slides[3].Hidden);
        Assert.False(pres.Slides[4].Hidden);
    }

    #endregion

    #region Hide Slides

    [SkippableFact]
    public void Execute_WithHiddenTrue_HidesAllSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Hidden);
        Assert.True(pres.Slides[1].Hidden);
        Assert.True(pres.Slides[2].Hidden);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithSpecificIndices_HidesOnlyThoseSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(5);
        var indices = JsonSerializer.Serialize(IndicesZeroTwoFour);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", true },
            { "slideIndices", indices }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.True(pres.Slides[0].Hidden);
        Assert.False(pres.Slides[1].Hidden);
        Assert.True(pres.Slides[2].Hidden);
        Assert.False(pres.Slides[3].Hidden);
        Assert.True(pres.Slides[4].Hidden);
        AssertModified(context);
    }

    #endregion

    #region Show Slides

    [SkippableFact]
    public void Execute_WithHiddenFalse_ShowsAllSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Hidden = true;
        pres.Slides[1].Hidden = true;
        pres.Slides[2].Hidden = true;
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", false }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        Assert.False(pres.Slides[0].Hidden);
        Assert.False(pres.Slides[1].Hidden);
        Assert.False(pres.Slides[2].Hidden);
        AssertModified(context);
    }

    [SkippableFact]
    public void Execute_WithSpecificIndices_ShowsOnlyThoseSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(5);
        pres.Slides[0].Hidden = true;
        pres.Slides[1].Hidden = true;
        pres.Slides[2].Hidden = true;
        pres.Slides[3].Hidden = true;
        pres.Slides[4].Hidden = true;
        var indices = JsonSerializer.Serialize(IndicesOneThree);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", false },
            { "slideIndices", indices }
        });

        _handler.Execute(context, parameters);

        Assert.True(pres.Slides[0].Hidden);
        Assert.False(pres.Slides[1].Hidden);
        Assert.True(pres.Slides[2].Hidden);
        Assert.False(pres.Slides[3].Hidden);
        Assert.True(pres.Slides[4].Hidden);
    }

    #endregion

    #region Default Behavior

    [SkippableFact]
    public void Execute_WithoutHiddenParam_DefaultsToFalse()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        pres.Slides[0].Hidden = true;
        pres.Slides[1].Hidden = true;
        var context = CreateContext(pres);
        var parameters = CreateEmptyParameters();

        _handler.Execute(context, parameters);

        Assert.False(pres.Slides[0].Hidden);
        Assert.False(pres.Slides[1].Hidden);
    }

    [SkippableFact]
    public void Execute_WithoutSlideIndices_AffectsAllSlides()
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(5);
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", true }
        });

        var res = _handler.Execute(context, parameters);

        Assert.IsType<SuccessResult>(res);
        foreach (var slide in pres.Slides)
            Assert.True(slide.Hidden);
        AssertModified(context);
    }

    #endregion

    #region Error Handling - Invalid Index

    [SkippableTheory]
    [InlineData(3, 3)]
    [InlineData(3, 5)]
    [InlineData(3, 100)]
    public void Execute_WithIndexOutOfRange_ThrowsArgumentException(int totalSlides, int invalidIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(totalSlides);
        var indices = JsonSerializer.Serialize(new[] { invalidIndex });
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", true },
            { "slideIndices", indices }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [SkippableTheory]
    [InlineData(-1)]
    [InlineData(-5)]
    public void Execute_WithNegativeIndex_ThrowsArgumentException(int negativeIndex)
    {
        SkipIfNotWindows();
        var pres = CreatePresentationWithSlides(3);
        var indices = JsonSerializer.Serialize(new[] { negativeIndex });
        var context = CreateContext(pres);
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "hidden", true },
            { "slideIndices", indices }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
        Assert.Contains("out of range", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion
}
