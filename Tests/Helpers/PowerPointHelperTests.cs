using Aspose.Slides;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Helpers;

/// <summary>
///     Unit tests for PowerPointHelper class
/// </summary>
public class PowerPointHelperTests : TestBase
{
    #region CountShapeTypes Tests

    [Fact]
    public void CountShapeTypes_WithTable_ShouldIncrementTablesCount()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50], [30]);

        int images = 0, tables = 0, charts = 0, smartArt = 0, audio = 0, video = 0;
        PowerPointHelper.CountShapeTypes(table, ref images, ref tables, ref charts, ref smartArt, ref audio, ref video);

        Assert.Equal(1, tables);
        Assert.Equal(0, images);
    }

    #endregion

    #region GenerateThumbnail Tests

    [Fact]
    public void GenerateThumbnail_ShouldReturnBase64String()
    {
        // Skip in evaluation mode as thumbnails may have watermarks
        if (IsEvaluationMode())
            return;

        using var presentation = new Presentation();
        var slide = presentation.Slides[0];

        var thumbnail = PowerPointHelper.GenerateThumbnail(slide);

        Assert.NotNull(thumbnail);
        Assert.NotEmpty(thumbnail);
        var bytes = Convert.FromBase64String(thumbnail);
        Assert.True(bytes.Length > 0);
    }

    #endregion

    #region ValidateSlideIndex Tests

    [Fact]
    public void ValidateSlideIndex_WithValidIndex_ShouldNotThrow()
    {
        using var presentation = new Presentation();

        var exception = Record.Exception(() =>
            PowerPointHelper.ValidateSlideIndex(0, presentation));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateSlideIndex_WithNegativeIndex_ShouldThrow()
    {
        using var presentation = new Presentation();

        var ex = Assert.Throws<ArgumentException>(() =>
            PowerPointHelper.ValidateSlideIndex(-1, presentation));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void ValidateSlideIndex_WithIndexExceedingCount_ShouldThrow()
    {
        using var presentation = new Presentation();

        var ex = Assert.Throws<ArgumentException>(() =>
            PowerPointHelper.ValidateSlideIndex(5, presentation));

        Assert.Contains("out of range", ex.Message);
        Assert.Contains("1 slides", ex.Message);
    }

    #endregion

    #region GetSlide Tests

    [Fact]
    public void GetSlide_WithValidIndex_ShouldReturnSlide()
    {
        using var presentation = new Presentation();

        var slide = PowerPointHelper.GetSlide(presentation, 0);

        Assert.NotNull(slide);
    }

    [Fact]
    public void GetSlide_WithInvalidIndex_ShouldThrow()
    {
        using var presentation = new Presentation();

        Assert.Throws<ArgumentException>(() =>
            PowerPointHelper.GetSlide(presentation, 5));
    }

    #endregion

    #region ValidateShapeIndex Tests

    [Fact]
    public void ValidateShapeIndex_WithNoShapes_ShouldThrow()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];

        var ex = Assert.Throws<ArgumentException>(() =>
            PowerPointHelper.ValidateShapeIndex(0, slide));

        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void ValidateShapeIndex_WithValidIndex_ShouldNotThrow()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var exception = Record.Exception(() =>
            PowerPointHelper.ValidateShapeIndex(0, slide));

        Assert.Null(exception);
    }

    #endregion

    #region GetShape Tests

    [Fact]
    public void GetShape_WithValidIndex_ShouldReturnShape()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

        var shape = PowerPointHelper.GetShape(slide, 0);

        Assert.NotNull(shape);
    }

    [Fact]
    public void GetShape_WithInvalidIndex_ShouldThrow()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];

        Assert.Throws<ArgumentException>(() =>
            PowerPointHelper.GetShape(slide, 0));
    }

    #endregion

    #region ValidateCollectionIndex Tests

    [Fact]
    public void ValidateCollectionIndex_Generic_WithValidIndex_ShouldNotThrow()
    {
        List<int> list = [1, 2, 3];

        var exception = Record.Exception(() =>
            PowerPointHelper.ValidateCollectionIndex(1, list));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateCollectionIndex_Generic_WithInvalidIndex_ShouldThrow()
    {
        List<int> list = [1, 2, 3];

        var ex = Assert.Throws<ArgumentException>(() =>
            PowerPointHelper.ValidateCollectionIndex(5, list, "Element"));

        Assert.Contains("Element", ex.Message);
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void ValidateCollectionIndex_Count_WithValidIndex_ShouldNotThrow()
    {
        var exception = Record.Exception(() =>
            PowerPointHelper.ValidateCollectionIndex(2, 5, "Row"));

        Assert.Null(exception);
    }

    [Fact]
    public void ValidateCollectionIndex_Count_WithInvalidIndex_ShouldThrow()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            PowerPointHelper.ValidateCollectionIndex(10, 5, "Column"));

        Assert.Contains("Column", ex.Message);
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region ExtractTextFromShape Tests

    [Fact]
    public void ExtractTextFromShape_WithAutoShape_ShouldExtractText()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
        shape.TextFrame.Text = "Hello World";

        List<string> textContent = [];
        PowerPointHelper.ExtractTextFromShape(shape, textContent);

        Assert.Single(textContent);
        // In evaluation mode, text may be truncated
        Assert.True(textContent[0].StartsWith("Hello"),
            $"Expected text to start with 'Hello', got: '{textContent[0]}'");
    }

    [Fact]
    public void ExtractTextFromShape_WithEmptyAutoShape_ShouldNotAddText()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);

        List<string> textContent = [];
        PowerPointHelper.ExtractTextFromShape(shape, textContent);

        Assert.Empty(textContent);
    }

    [Fact]
    public void ExtractTextFromShape_WithTable_ShouldExtractCellText()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(10, 10, [50, 50], [30, 30]);
        table[0, 0].TextFrame.Text = "Cell 1";
        table[1, 0].TextFrame.Text = "Cell 2";

        List<string> textContent = [];
        PowerPointHelper.ExtractTextFromShape(table, textContent);

        Assert.Equal(2, textContent.Count);
        // In evaluation mode, text may be truncated
        Assert.True(textContent.Any(t => t.StartsWith("Cell")),
            $"Expected at least one cell text to start with 'Cell', got: {string.Join(", ", textContent)}");
    }

    #endregion

    #region CountTextCharacters Tests

    [Fact]
    public void CountTextCharacters_WithAutoShape_ShouldCountCorrectly()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);
        shape.TextFrame.Text = "Hello";

        var count = PowerPointHelper.CountTextCharacters(shape);

        Assert.Equal(5, count);
    }

    [Fact]
    public void CountTextCharacters_WithEmptyShape_ShouldReturnZero()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 50);

        var count = PowerPointHelper.CountTextCharacters(shape);

        Assert.Equal(0, count);
    }

    #endregion
}
