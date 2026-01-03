using System.Drawing;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Tests.Core.Helpers;

/// <summary>
///     Unit tests for ColorHelper class
/// </summary>
public class ColorHelperTests
{
    #region ParseColor (default) Tests

    [Fact]
    public void ParseColor_WithNullOrEmpty_ShouldReturnBlack()
    {
        Assert.Equal(Color.Black, ColorHelper.ParseColor(null));
        Assert.Equal(Color.Black, ColorHelper.ParseColor(""));
        Assert.Equal(Color.Black, ColorHelper.ParseColor("   "));
    }

    [Fact]
    public void ParseColor_WithHexRgb_ShouldParseCorrectly()
    {
        var result = ColorHelper.ParseColor("#FF0000");

        Assert.Equal(255, result.R);
        Assert.Equal(0, result.G);
        Assert.Equal(0, result.B);
    }

    [Fact]
    public void ParseColor_WithHexRgbNoHash_ShouldParseCorrectly()
    {
        var result = ColorHelper.ParseColor("00FF00");

        Assert.Equal(0, result.R);
        Assert.Equal(255, result.G);
        Assert.Equal(0, result.B);
    }

    [Fact]
    public void ParseColor_WithHexArgb_ShouldParseCorrectly()
    {
        var result = ColorHelper.ParseColor("#80FF0000");

        Assert.Equal(128, result.A);
        Assert.Equal(255, result.R);
        Assert.Equal(0, result.G);
        Assert.Equal(0, result.B);
    }

    [Fact]
    public void ParseColor_WithRgbCommas_ShouldParseCorrectly()
    {
        var result = ColorHelper.ParseColor("255,128,64");

        Assert.Equal(255, result.R);
        Assert.Equal(128, result.G);
        Assert.Equal(64, result.B);
    }

    [Fact]
    public void ParseColor_WithRgbCommasAndSpaces_ShouldParseCorrectly()
    {
        var result = ColorHelper.ParseColor("255, 128, 64");

        Assert.Equal(255, result.R);
        Assert.Equal(128, result.G);
        Assert.Equal(64, result.B);
    }

    [Fact]
    public void ParseColor_WithRgbOutOfRange_ShouldClamp()
    {
        var result = ColorHelper.ParseColor("300, -50, 128");

        Assert.Equal(255, result.R);
        Assert.Equal(0, result.G);
        Assert.Equal(128, result.B);
    }

    [Fact]
    public void ParseColor_WithNamedColor_ShouldParseCorrectly()
    {
        var result = ColorHelper.ParseColor("Red");

        Assert.Equal(Color.Red.R, result.R);
        Assert.Equal(Color.Red.G, result.G);
        Assert.Equal(Color.Red.B, result.B);
    }

    [Fact]
    public void ParseColor_WithInvalidString_ShouldReturnBlack()
    {
        var result = ColorHelper.ParseColor("not_a_color");

        Assert.Equal(Color.Black, result);
    }

    #endregion

    #region ParseColor (with default color) Tests

    [Fact]
    public void ParseColor_WithDefaultColor_NullInput_ShouldReturnDefault()
    {
        var defaultColor = Color.Blue;

        var result = ColorHelper.ParseColor(null, defaultColor);

        Assert.Equal(defaultColor, result);
    }

    [Fact]
    public void ParseColor_WithDefaultColor_ValidInput_ShouldParseCorrectly()
    {
        var defaultColor = Color.Blue;

        var result = ColorHelper.ParseColor("#FF0000", defaultColor);

        Assert.Equal(255, result.R);
        Assert.Equal(0, result.G);
        Assert.Equal(0, result.B);
    }

    [Fact]
    public void ParseColor_WithDefaultColor_InvalidInput_ShouldReturnDefault()
    {
        var defaultColor = Color.Yellow;

        var result = ColorHelper.ParseColor("invalid", defaultColor);

        Assert.Equal(defaultColor, result);
    }

    #endregion

    #region ParseColor (with throwOnError) Tests

    [Fact]
    public void ParseColor_ThrowOnError_WithNull_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ColorHelper.ParseColor(null, true));

        Assert.Contains("cannot be null or empty", ex.Message);
    }

    [Fact]
    public void ParseColor_ThrowOnError_WithEmpty_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ColorHelper.ParseColor("", true));

        Assert.Contains("cannot be null or empty", ex.Message);
    }

    [Fact]
    public void ParseColor_ThrowOnError_WithInvalid_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            ColorHelper.ParseColor("invalid_color", true));

        Assert.Contains("Unable to parse color", ex.Message);
    }

    [Fact]
    public void ParseColor_ThrowOnError_WithValid_ShouldReturnColor()
    {
        var result = ColorHelper.ParseColor("#FF0000", true);

        Assert.Equal(255, result.R);
        Assert.Equal(0, result.G);
        Assert.Equal(0, result.B);
    }

    [Fact]
    public void ParseColor_ThrowOnErrorFalse_WithInvalid_ShouldReturnBlack()
    {
        var result = ColorHelper.ParseColor("invalid", false);

        Assert.Equal(Color.Black, result);
    }

    #endregion

    #region TryParseColor Tests

    [Fact]
    public void TryParseColor_WithValidHex_ShouldReturnTrue()
    {
        var success = ColorHelper.TryParseColor("#FF0000", out var color);

        Assert.True(success);
        Assert.Equal(255, color.R);
        Assert.Equal(0, color.G);
        Assert.Equal(0, color.B);
    }

    [Fact]
    public void TryParseColor_WithValidRgb_ShouldReturnTrue()
    {
        var success = ColorHelper.TryParseColor("128,64,32", out var color);

        Assert.True(success);
        Assert.Equal(128, color.R);
        Assert.Equal(64, color.G);
        Assert.Equal(32, color.B);
    }

    [Fact]
    public void TryParseColor_WithValidName_ShouldReturnTrue()
    {
        var success = ColorHelper.TryParseColor("Blue", out var color);

        Assert.True(success);
        Assert.Equal(Color.Blue.R, color.R);
        Assert.Equal(Color.Blue.G, color.G);
        Assert.Equal(Color.Blue.B, color.B);
    }

    [Fact]
    public void TryParseColor_WithNull_ShouldReturnFalse()
    {
        var success = ColorHelper.TryParseColor(null, out var color);

        Assert.False(success);
        Assert.Equal(Color.Black, color);
    }

    [Fact]
    public void TryParseColor_WithEmpty_ShouldReturnFalse()
    {
        var success = ColorHelper.TryParseColor("", out var color);

        Assert.False(success);
        Assert.Equal(Color.Black, color);
    }

    [Fact]
    public void TryParseColor_WithInvalid_ShouldReturnFalse()
    {
        var success = ColorHelper.TryParseColor("not_a_color", out var color);

        Assert.False(success);
        Assert.Equal(Color.Black, color);
    }

    #endregion

    #region ToPdfColor Tests

    [Fact]
    public void ToPdfColor_WithRed_ShouldConvertCorrectly()
    {
        var color = Color.FromArgb(255, 0, 0);

        var pdfColor = ColorHelper.ToPdfColor(color);

        Assert.NotNull(pdfColor);
    }

    [Fact]
    public void ToPdfColor_WithBlack_ShouldConvertCorrectly()
    {
        var color = Color.FromArgb(0, 0, 0);

        var pdfColor = ColorHelper.ToPdfColor(color);

        Assert.NotNull(pdfColor);
    }

    [Fact]
    public void ToPdfColor_WithWhite_ShouldConvertCorrectly()
    {
        var color = Color.FromArgb(255, 255, 255);

        var pdfColor = ColorHelper.ToPdfColor(color);

        Assert.NotNull(pdfColor);
    }

    [Fact]
    public void ToPdfColor_WithMixedColor_ShouldConvertCorrectly()
    {
        var color = Color.FromArgb(128, 64, 32);

        var pdfColor = ColorHelper.ToPdfColor(color);

        Assert.NotNull(pdfColor);
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void ParseColor_WithWhitespace_ShouldTrim()
    {
        var result = ColorHelper.ParseColor("  #FF0000  ");

        Assert.Equal(255, result.R);
        Assert.Equal(0, result.G);
        Assert.Equal(0, result.B);
    }

    [Fact]
    public void ParseColor_WithLowercaseHex_ShouldParseCorrectly()
    {
        var result = ColorHelper.ParseColor("#ff00ff");

        Assert.Equal(255, result.R);
        Assert.Equal(0, result.G);
        Assert.Equal(255, result.B);
    }

    [Fact]
    public void ParseColor_WithMixedCaseHex_ShouldParseCorrectly()
    {
        var result = ColorHelper.ParseColor("#Ff00fF");

        Assert.Equal(255, result.R);
        Assert.Equal(0, result.G);
        Assert.Equal(255, result.B);
    }

    [Fact]
    public void ParseColor_WithPartialRgb_ShouldReturnBlack()
    {
        var result = ColorHelper.ParseColor("255,128");

        Assert.Equal(Color.Black, result);
    }

    [Fact]
    public void ParseColor_WithTooManyRgbValues_ShouldReturnBlack()
    {
        var result = ColorHelper.ParseColor("255,128,64,32");

        Assert.Equal(Color.Black, result);
    }

    [Fact]
    public void ParseColor_WithInvalidHexLength_ShouldReturnBlack()
    {
        var result = ColorHelper.ParseColor("#FFF");

        Assert.Equal(Color.Black, result);
    }

    #endregion
}