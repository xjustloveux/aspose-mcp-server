using Aspose.Pdf;
using AsposeMcpServer.Handlers.Pdf.Watermark;

namespace AsposeMcpServer.Tests.Handlers.Pdf.Watermark;

public class PdfWatermarkHelperTests
{
    #region ParseColor Tests - Named Colors

    [Theory]
    [InlineData("red")]
    [InlineData("Red")]
    [InlineData("RED")]
    public void ParseColor_WithRed_ReturnsRed(string input)
    {
        var result = PdfWatermarkHelper.ParseColor(input);

        Assert.Equal(Color.Red, result);
    }

    [Theory]
    [InlineData("blue")]
    [InlineData("green")]
    [InlineData("black")]
    [InlineData("white")]
    [InlineData("yellow")]
    [InlineData("orange")]
    [InlineData("purple")]
    [InlineData("pink")]
    [InlineData("cyan")]
    [InlineData("magenta")]
    [InlineData("lightgray")]
    [InlineData("darkgray")]
    public void ParseColor_WithNamedColors_ReturnsCorrectColor(string input)
    {
        var result = PdfWatermarkHelper.ParseColor(input);

        Assert.NotNull(result);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("unknown")]
    [InlineData("")]
    public void ParseColor_WithInvalidOrEmpty_ReturnsGray(string input)
    {
        var result = PdfWatermarkHelper.ParseColor(input);

        Assert.Equal(Color.Gray, result);
    }

    [Fact]
    public void ParseColor_WithNull_ReturnsGray()
    {
        var result = PdfWatermarkHelper.ParseColor(null!);

        Assert.Equal(Color.Gray, result);
    }

    #endregion

    #region ParseColor Tests - Hex Colors

    [Fact]
    public void ParseColor_WithValidHexColor_ReturnsColor()
    {
        var result = PdfWatermarkHelper.ParseColor("#FF0000");

        Assert.NotNull(result);
    }

    [Fact]
    public void ParseColor_WithValidHexColorWithAlpha_ReturnsColor()
    {
        var result = PdfWatermarkHelper.ParseColor("#FF0000FF");

        Assert.NotNull(result);
    }

    [Theory]
    [InlineData("#FFF")]
    [InlineData("#FFFFF")]
    [InlineData("#FFFFFFFFFF")]
    public void ParseColor_WithInvalidHexLength_ReturnsGray(string input)
    {
        var result = PdfWatermarkHelper.ParseColor(input);

        Assert.Equal(Color.Gray, result);
    }

    [Fact]
    public void ParseColor_WithInvalidHexCharacters_ReturnsGray()
    {
        var result = PdfWatermarkHelper.ParseColor("#GGGGGG");

        Assert.Equal(Color.Gray, result);
    }

    #endregion

    #region ParsePageRange Tests - Empty/Null

    [Fact]
    public void ParsePageRange_WithNull_ReturnsAllPages()
    {
        var result = PdfWatermarkHelper.ParsePageRange(null, 5);

        Assert.Equal(5, result.Count);
        Assert.Equal([1, 2, 3, 4, 5], result);
    }

    [Fact]
    public void ParsePageRange_WithEmpty_ReturnsAllPages()
    {
        var result = PdfWatermarkHelper.ParsePageRange("", 5);

        Assert.Equal(5, result.Count);
    }

    #endregion

    #region ParsePageRange Tests - Single Pages

    [Fact]
    public void ParsePageRange_WithSinglePage_ReturnsSinglePage()
    {
        var result = PdfWatermarkHelper.ParsePageRange("3", 5);

        Assert.Single(result);
        Assert.Equal(3, result[0]);
    }

    [Fact]
    public void ParsePageRange_WithMultipleSinglePages_ReturnsAllPages()
    {
        var result = PdfWatermarkHelper.ParsePageRange("1,3,5", 5);

        Assert.Equal(3, result.Count);
        Assert.Contains(1, result);
        Assert.Contains(3, result);
        Assert.Contains(5, result);
    }

    [Fact]
    public void ParsePageRange_WithDuplicates_ReturnsDeduplicated()
    {
        var result = PdfWatermarkHelper.ParsePageRange("1,1,2,2", 5);

        Assert.Equal(2, result.Count);
    }

    #endregion

    #region ParsePageRange Tests - Ranges

    [Fact]
    public void ParsePageRange_WithRange_ReturnsAllInRange()
    {
        var result = PdfWatermarkHelper.ParsePageRange("2-4", 5);

        Assert.Equal(3, result.Count);
        Assert.Contains(2, result);
        Assert.Contains(3, result);
        Assert.Contains(4, result);
    }

    [Fact]
    public void ParsePageRange_WithMixedRangesAndPages_ReturnsAll()
    {
        var result = PdfWatermarkHelper.ParsePageRange("1,3-5", 5);

        Assert.Equal(4, result.Count);
        Assert.Contains(1, result);
        Assert.Contains(3, result);
        Assert.Contains(4, result);
        Assert.Contains(5, result);
    }

    [Fact]
    public void ParsePageRange_ReturnsOrdered()
    {
        var result = PdfWatermarkHelper.ParsePageRange("5,1,3", 5);

        Assert.Equal([1, 3, 5], result);
    }

    #endregion

    #region ParsePageRange Tests - Error Cases

    [Fact]
    public void ParsePageRange_WithInvalidNumber_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PdfWatermarkHelper.ParsePageRange("abc", 5));

        Assert.Contains("Invalid page number", ex.Message);
    }

    [Fact]
    public void ParsePageRange_WithPageZero_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PdfWatermarkHelper.ParsePageRange("0", 5));

        Assert.Contains("out of bounds", ex.Message);
    }

    [Fact]
    public void ParsePageRange_WithPageTooLarge_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PdfWatermarkHelper.ParsePageRange("10", 5));

        Assert.Contains("out of bounds", ex.Message);
    }

    [Fact]
    public void ParsePageRange_WithInvalidRangeFormat_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PdfWatermarkHelper.ParsePageRange("1-2-3", 5));

        Assert.Contains("Invalid page range format", ex.Message);
    }

    [Fact]
    public void ParsePageRange_WithReverseRange_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PdfWatermarkHelper.ParsePageRange("5-2", 5));

        Assert.Contains("out of bounds", ex.Message);
    }

    [Fact]
    public void ParsePageRange_WithRangeOutOfBounds_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PdfWatermarkHelper.ParsePageRange("1-10", 5));

        Assert.Contains("out of bounds", ex.Message);
    }

    #endregion
}
