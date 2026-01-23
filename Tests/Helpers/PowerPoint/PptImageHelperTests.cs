using Aspose.Slides;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers.PowerPoint;

public class PptImageHelperTests : PptTestBase
{
    #region GetPictureFrames Tests

    [Fact]
    public void GetPictureFrames_WithNoImages_ReturnsEmptyList()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];

        var result = PptImageHelper.GetPictureFrames(slide);

        Assert.Empty(result);
    }

    #endregion

    #region CalculateDimensions Tests

    [Fact]
    public void CalculateDimensions_WithBothWidthAndHeight_ReturnsBoth()
    {
        var result = PptImageHelper.CalculateDimensions(200, 100, 400, 300);

        Assert.Equal(200, result.width);
        Assert.Equal(100, result.height);
    }

    [Fact]
    public void CalculateDimensions_WithOnlyWidth_CalculatesHeight()
    {
        var result = PptImageHelper.CalculateDimensions(200, null, 400, 300);

        Assert.Equal(200, result.width);
        Assert.Equal(150, result.height);
    }

    [Fact]
    public void CalculateDimensions_WithOnlyHeight_CalculatesWidth()
    {
        var result = PptImageHelper.CalculateDimensions(null, 150, 400, 300);

        Assert.Equal(200, result.width);
        Assert.Equal(150, result.height);
    }

    [Fact]
    public void CalculateDimensions_WithNoWidthOrHeight_ReturnsDefault()
    {
        var result = PptImageHelper.CalculateDimensions(null, null, 400, 300);

        Assert.Equal(300, result.width);
        Assert.Equal(225, result.height);
    }

    [Fact]
    public void CalculateDimensions_WithZeroPixelWidth_HandlesGracefully()
    {
        var result = PptImageHelper.CalculateDimensions(200, null, 0, 300);

        Assert.Equal(200, result.width);
        Assert.Equal(200, result.height);
    }

    [Fact]
    public void CalculateDimensions_WithZeroPixelHeight_HandlesGracefully()
    {
        var result = PptImageHelper.CalculateDimensions(null, 150, 400, 0);

        Assert.Equal(150, result.width);
        Assert.Equal(150, result.height);
    }

    #endregion

    #region CalculateResizeSize Tests

    [Fact]
    public void CalculateResizeSize_WithNoConstraints_ReturnsSameSize()
    {
        var result = PptImageHelper.CalculateResizeSize(400, 300, null, null);

        Assert.Equal(400, result.Width);
        Assert.Equal(300, result.Height);
    }

    [Fact]
    public void CalculateResizeSize_WithMaxWidth_ScalesDown()
    {
        var result = PptImageHelper.CalculateResizeSize(400, 300, 200, null);

        Assert.Equal(200, result.Width);
        Assert.Equal(150, result.Height);
    }

    [Fact]
    public void CalculateResizeSize_WithMaxHeight_ScalesDown()
    {
        var result = PptImageHelper.CalculateResizeSize(400, 300, null, 150);

        Assert.Equal(200, result.Width);
        Assert.Equal(150, result.Height);
    }

    [Fact]
    public void CalculateResizeSize_WithBothConstraints_ScalesToFit()
    {
        var result = PptImageHelper.CalculateResizeSize(400, 300, 200, 200);

        Assert.True(result.Width <= 200);
        Assert.True(result.Height <= 200);
    }

    [Fact]
    public void CalculateResizeSize_WithImageSmallerThanMax_ReturnsSameSize()
    {
        var result = PptImageHelper.CalculateResizeSize(100, 100, 200, 200);

        Assert.Equal(100, result.Width);
        Assert.Equal(100, result.Height);
    }

    #endregion

    #region ParseSlideIndexes Tests

    [Fact]
    public void ParseSlideIndexes_WithNull_ReturnsAllSlides()
    {
        var result = PptImageHelper.ParseSlideIndexes(null, 5);

        Assert.Equal(5, result.Count);
        Assert.Equal([0, 1, 2, 3, 4], result);
    }

    [Fact]
    public void ParseSlideIndexes_WithEmpty_ReturnsAllSlides()
    {
        var result = PptImageHelper.ParseSlideIndexes("", 5);

        Assert.Equal(5, result.Count);
    }

    [Fact]
    public void ParseSlideIndexes_WithSingleIndex_ReturnsSingleSlide()
    {
        var result = PptImageHelper.ParseSlideIndexes("2", 5);

        Assert.Single(result);
        Assert.Equal(2, result[0]);
    }

    [Fact]
    public void ParseSlideIndexes_WithMultipleIndexes_ReturnsAllSpecified()
    {
        var result = PptImageHelper.ParseSlideIndexes("0,2,4", 5);

        Assert.Equal(3, result.Count);
        Assert.Contains(0, result);
        Assert.Contains(2, result);
        Assert.Contains(4, result);
    }

    [Fact]
    public void ParseSlideIndexes_WithDuplicates_ReturnsDeduplicated()
    {
        var result = PptImageHelper.ParseSlideIndexes("1,1,2,2", 5);

        Assert.Equal(2, result.Count);
    }

    [Fact]
    public void ParseSlideIndexes_WithInvalidIndex_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptImageHelper.ParseSlideIndexes("abc", 5));

        Assert.Contains("Invalid slide index", ex.Message);
    }

    [Fact]
    public void ParseSlideIndexes_WithOutOfRangeIndex_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptImageHelper.ParseSlideIndexes("10", 5));

        Assert.Contains("must be between 0 and", ex.Message);
    }

    [Fact]
    public void ParseSlideIndexes_WithNegativeIndex_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptImageHelper.ParseSlideIndexes("-1", 5));

        Assert.Contains("must be between 0 and", ex.Message);
    }

    #endregion

    #region ComputeImageHash Tests

    [Fact]
    public void ComputeImageHash_WithSameData_ReturnsSameHash()
    {
        var data = new byte[] { 1, 2, 3, 4, 5 };

        var result1 = PptImageHelper.ComputeImageHash(data);
        var result2 = PptImageHelper.ComputeImageHash(data);

        Assert.Equal(result1, result2);
    }

    [Fact]
    public void ComputeImageHash_WithDifferentData_ReturnsDifferentHash()
    {
        var data1 = new byte[] { 1, 2, 3, 4, 5 };
        var data2 = new byte[] { 5, 4, 3, 2, 1 };

        var result1 = PptImageHelper.ComputeImageHash(data1);
        var result2 = PptImageHelper.ComputeImageHash(data2);

        Assert.NotEqual(result1, result2);
    }

    [Fact]
    public void ComputeImageHash_WithEmptyData_ReturnsHash()
    {
        var data = Array.Empty<byte>();

        var result = PptImageHelper.ComputeImageHash(data);

        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    #endregion

    #region ValidateImageIndex Tests

    [Fact]
    public void ValidateImageIndex_WithValidIndex_DoesNotThrow()
    {
        var exception = Record.Exception(() =>
            PptImageHelper.ValidateImageIndex(0, 0, 5));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateImageIndex_WithNegativeIndex_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptImageHelper.ValidateImageIndex(-1, 0, 5));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void ValidateImageIndex_WithIndexTooLarge_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptImageHelper.ValidateImageIndex(10, 0, 5));

        Assert.Contains("out of range", ex.Message);
        Assert.Contains("5 image(s)", ex.Message);
    }

    #endregion
}
