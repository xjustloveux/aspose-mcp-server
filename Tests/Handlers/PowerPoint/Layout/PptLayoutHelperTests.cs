using Aspose.Slides;
using AsposeMcpServer.Handlers.PowerPoint.Layout;

namespace AsposeMcpServer.Tests.Handlers.PowerPoint.Layout;

public class PptLayoutHelperTests
{
    #region BuildLayoutsList Tests

    [Fact]
    public void BuildLayoutsList_ReturnsLayoutInfo()
    {
        using var presentation = new Presentation();
        var layouts = presentation.Masters[0].LayoutSlides;

        var result = PptLayoutHelper.BuildLayoutsList(layouts);

        Assert.NotEmpty(result);
        Assert.All(result, item =>
        {
            Assert.NotNull(item);
            var type = item.GetType();
            Assert.NotNull(type.GetProperty("index"));
            Assert.NotNull(type.GetProperty("name"));
            Assert.NotNull(type.GetProperty("layoutType"));
        });
    }

    #endregion

    #region SupportedLayoutTypes Tests

    [Fact]
    public void SupportedLayoutTypes_ContainsCommonLayouts()
    {
        var types = PptLayoutHelper.SupportedLayoutTypes;

        Assert.Contains("title", types);
        Assert.Contains("blank", types);
        Assert.Contains("titleonly", types);
    }

    #endregion

    #region FindLayoutByType Tests

    [Fact]
    public void FindLayoutByType_WithBlankLayout_ReturnsLayout()
    {
        using var presentation = new Presentation();

        var result = PptLayoutHelper.FindLayoutByType(presentation, "blank");

        Assert.NotNull(result);
        Assert.Equal(SlideLayoutType.Blank, result.LayoutType);
    }

    [Fact]
    public void FindLayoutByType_WithTitleLayout_ReturnsLayout()
    {
        using var presentation = new Presentation();

        var result = PptLayoutHelper.FindLayoutByType(presentation, "title");

        Assert.NotNull(result);
        Assert.Equal(SlideLayoutType.Title, result.LayoutType);
    }

    [Fact]
    public void FindLayoutByType_CaseInsensitive_ReturnsLayout()
    {
        using var presentation = new Presentation();

        var result1 = PptLayoutHelper.FindLayoutByType(presentation, "BLANK");
        var result2 = PptLayoutHelper.FindLayoutByType(presentation, "Blank");
        var result3 = PptLayoutHelper.FindLayoutByType(presentation, "blank");

        Assert.Equal(result1.LayoutType, result2.LayoutType);
        Assert.Equal(result2.LayoutType, result3.LayoutType);
    }

    [Fact]
    public void FindLayoutByType_WithUnknownType_ThrowsArgumentException()
    {
        using var presentation = new Presentation();

        var ex = Assert.Throws<ArgumentException>(() =>
            PptLayoutHelper.FindLayoutByType(presentation, "unknownLayout"));

        Assert.Contains("Unknown layout type", ex.Message);
        Assert.Contains("Supported types", ex.Message);
    }

    [Theory]
    [InlineData("titleonly")]
    [InlineData("sectionheader")]
    public void FindLayoutByType_WithValidTypes_ReturnsLayout(string layoutType)
    {
        using var presentation = new Presentation();

        var result = PptLayoutHelper.FindLayoutByType(presentation, layoutType);

        Assert.NotNull(result);
    }

    #endregion

    #region ValidateSlideIndices Tests

    [Fact]
    public void ValidateSlideIndices_WithValidIndices_DoesNotThrow()
    {
        var exception = Record.Exception(() =>
            PptLayoutHelper.ValidateSlideIndices([0, 1, 2], 5));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateSlideIndices_WithNegativeIndex_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptLayoutHelper.ValidateSlideIndices([0, -1, 2], 5));

        Assert.Contains("Invalid slide indices", ex.Message);
        Assert.Contains("-1", ex.Message);
    }

    [Fact]
    public void ValidateSlideIndices_WithIndexTooLarge_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptLayoutHelper.ValidateSlideIndices([0, 1, 10], 5));

        Assert.Contains("Invalid slide indices", ex.Message);
        Assert.Contains("10", ex.Message);
        Assert.Contains("Valid range: 0 to 4", ex.Message);
    }

    [Fact]
    public void ValidateSlideIndices_WithMultipleInvalidIndices_ReportsAll()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptLayoutHelper.ValidateSlideIndices([-1, 10, 20], 5));

        Assert.Contains("-1", ex.Message);
        Assert.Contains("10", ex.Message);
        Assert.Contains("20", ex.Message);
    }

    [Fact]
    public void ValidateSlideIndices_WithEmptyArray_DoesNotThrow()
    {
        var exception = Record.Exception(() =>
            PptLayoutHelper.ValidateSlideIndices([], 5));

        Assert.Null(exception);
    }

    #endregion

    #region ParseSlideIndicesJson Tests

    [Fact]
    public void ParseSlideIndicesJson_WithNull_ReturnsNull()
    {
        var result = PptLayoutHelper.ParseSlideIndicesJson(null);

        Assert.Null(result);
    }

    [Fact]
    public void ParseSlideIndicesJson_WithEmpty_ReturnsNull()
    {
        var result = PptLayoutHelper.ParseSlideIndicesJson("");

        Assert.Null(result);
    }

    [Fact]
    public void ParseSlideIndicesJson_WithWhitespace_ReturnsNull()
    {
        var result = PptLayoutHelper.ParseSlideIndicesJson("   ");

        Assert.Null(result);
    }

    [Fact]
    public void ParseSlideIndicesJson_WithValidArray_ReturnsIndices()
    {
        var result = PptLayoutHelper.ParseSlideIndicesJson("[0, 1, 2]");

        Assert.NotNull(result);
        Assert.Equal(3, result.Length);
        Assert.Equal(0, result[0]);
        Assert.Equal(1, result[1]);
        Assert.Equal(2, result[2]);
    }

    [Fact]
    public void ParseSlideIndicesJson_WithSingleElement_ReturnsIndices()
    {
        var result = PptLayoutHelper.ParseSlideIndicesJson("[5]");

        Assert.NotNull(result);
        Assert.Single(result);
        Assert.Equal(5, result[0]);
    }

    [Fact]
    public void ParseSlideIndicesJson_WithEmptyArray_ReturnsEmptyArray()
    {
        var result = PptLayoutHelper.ParseSlideIndicesJson("[]");

        Assert.NotNull(result);
        Assert.Empty(result);
    }

    [Fact]
    public void ParseSlideIndicesJson_WithInvalidJson_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptLayoutHelper.ParseSlideIndicesJson("not json"));

        Assert.Contains("Invalid slideIndices format", ex.Message);
        Assert.Contains("Expected JSON array", ex.Message);
    }

    [Fact]
    public void ParseSlideIndicesJson_WithObjectInsteadOfArray_ThrowsArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PptLayoutHelper.ParseSlideIndicesJson("{\"key\": 1}"));

        Assert.Contains("Invalid slideIndices format", ex.Message);
    }

    #endregion
}
