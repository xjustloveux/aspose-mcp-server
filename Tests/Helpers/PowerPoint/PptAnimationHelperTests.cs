using Aspose.Slides.Animation;
using AsposeMcpServer.Helpers.PowerPoint;

namespace AsposeMcpServer.Tests.Helpers.PowerPoint;

public class PptAnimationHelperTests
{
    #region ParseEffectType Tests

    [Fact]
    public void ParseEffectType_WithNull_ReturnsFade()
    {
        var result = PptAnimationHelper.ParseEffectType(null);

        Assert.Equal(EffectType.Fade, result);
    }

    [Fact]
    public void ParseEffectType_WithEmpty_ReturnsFade()
    {
        var result = PptAnimationHelper.ParseEffectType("");

        Assert.Equal(EffectType.Fade, result);
    }

    [Theory]
    [InlineData("Fade", EffectType.Fade)]
    [InlineData("fade", EffectType.Fade)]
    [InlineData("FADE", EffectType.Fade)]
    [InlineData("Fly", EffectType.Fly)]
    [InlineData("fly", EffectType.Fly)]
    [InlineData("Wipe", EffectType.Wipe)]
    [InlineData("Zoom", EffectType.Zoom)]
    [InlineData("Appear", EffectType.Appear)]
    public void ParseEffectType_WithValidValues_ReturnsCorrectType(string input, EffectType expected)
    {
        var result = PptAnimationHelper.ParseEffectType(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("notaneffect")]
    public void ParseEffectType_WithInvalidValues_ReturnsFade(string input)
    {
        var result = PptAnimationHelper.ParseEffectType(input);

        Assert.Equal(EffectType.Fade, result);
    }

    #endregion

    #region ParseEffectSubtype Tests

    [Fact]
    public void ParseEffectSubtype_WithNull_ReturnsNone()
    {
        var result = PptAnimationHelper.ParseEffectSubtype(null);

        Assert.Equal(EffectSubtype.None, result);
    }

    [Fact]
    public void ParseEffectSubtype_WithEmpty_ReturnsNone()
    {
        var result = PptAnimationHelper.ParseEffectSubtype("");

        Assert.Equal(EffectSubtype.None, result);
    }

    [Theory]
    [InlineData("None", EffectSubtype.None)]
    [InlineData("none", EffectSubtype.None)]
    [InlineData("Left", EffectSubtype.Left)]
    [InlineData("left", EffectSubtype.Left)]
    [InlineData("Right", EffectSubtype.Right)]
    [InlineData("Top", EffectSubtype.Top)]
    [InlineData("Bottom", EffectSubtype.Bottom)]
    public void ParseEffectSubtype_WithValidValues_ReturnsCorrectSubtype(string input, EffectSubtype expected)
    {
        var result = PptAnimationHelper.ParseEffectSubtype(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    public void ParseEffectSubtype_WithInvalidValues_ReturnsNone(string input)
    {
        var result = PptAnimationHelper.ParseEffectSubtype(input);

        Assert.Equal(EffectSubtype.None, result);
    }

    #endregion

    #region ParseTriggerType Tests

    [Fact]
    public void ParseTriggerType_WithNull_ReturnsOnClick()
    {
        var result = PptAnimationHelper.ParseTriggerType(null);

        Assert.Equal(EffectTriggerType.OnClick, result);
    }

    [Fact]
    public void ParseTriggerType_WithEmpty_ReturnsOnClick()
    {
        var result = PptAnimationHelper.ParseTriggerType("");

        Assert.Equal(EffectTriggerType.OnClick, result);
    }

    [Theory]
    [InlineData("OnClick", EffectTriggerType.OnClick)]
    [InlineData("onclick", EffectTriggerType.OnClick)]
    [InlineData("ONCLICK", EffectTriggerType.OnClick)]
    [InlineData("WithPrevious", EffectTriggerType.WithPrevious)]
    [InlineData("withprevious", EffectTriggerType.WithPrevious)]
    [InlineData("AfterPrevious", EffectTriggerType.AfterPrevious)]
    [InlineData("afterprevious", EffectTriggerType.AfterPrevious)]
    public void ParseTriggerType_WithValidValues_ReturnsCorrectTrigger(string input, EffectTriggerType expected)
    {
        var result = PptAnimationHelper.ParseTriggerType(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("click")]
    public void ParseTriggerType_WithInvalidValues_ReturnsOnClick(string input)
    {
        var result = PptAnimationHelper.ParseTriggerType(input);

        Assert.Equal(EffectTriggerType.OnClick, result);
    }

    #endregion
}
