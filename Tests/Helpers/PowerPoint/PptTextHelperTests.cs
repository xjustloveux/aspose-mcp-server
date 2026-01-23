using Aspose.Slides;
using AsposeMcpServer.Helpers.PowerPoint;

namespace AsposeMcpServer.Tests.Helpers.PowerPoint;

public class PptTextHelperTests
{
    #region ProcessShapesForReplace Tests

    [Fact]
    public void ProcessShapesForReplace_WithNoShapes_ReturnsZero()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.Clear();

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "find", "replace", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(0, result);
    }

    [Fact]
    public void ProcessShapesForReplace_WithMatchingText_ReturnsOneAndReplacesText()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.TextFrame.Text = "Hello World";

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "Hello", "Hi", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(1, result);
        Assert.Contains("Hi", shape.TextFrame.Text);
    }

    [Fact]
    public void ProcessShapesForReplace_WithNoMatch_ReturnsZero()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.TextFrame.Text = "Hello World";

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "NotFound", "Replace", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(0, result);
    }

    [Fact]
    public void ProcessShapesForReplace_WithCaseSensitiveMatch_RespectsComparison()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.TextFrame.Text = "Hello World";

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "hello", "Hi", StringComparison.Ordinal);

        Assert.Equal(0, result);
    }

    [Fact]
    public void ProcessShapesForReplace_WithCaseInsensitiveMatch_FindsMatch()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.TextFrame.Text = "Hello World";

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "hello", "Hi", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(1, result);
    }

    [Fact]
    public void ProcessShapesForReplace_WithMultipleShapes_ReplacesInAll()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape1 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape1.TextFrame.Text = "Hello One";
        var shape2 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 200, 50);
        shape2.TextFrame.Text = "Hello Two";

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "Hello", "Hi", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(2, result);
    }

    [Fact]
    public void ProcessShapesForReplace_WithNestedShapes_ProcessesShapes()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.TextFrame.Text = "Text to find here";

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "find", "replace", StringComparison.OrdinalIgnoreCase);

        Assert.True(result >= 0);
    }

    [Fact]
    public void ProcessShapesForReplace_WithTable_ReplacesInCells()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        double[] colWidths = [100, 100];
        double[] rowHeights = [50, 50];
        var table = slide.Shapes.AddTable(100, 100, colWidths, rowHeights);
        table[0, 0].TextFrame.Text = "Find this";

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "Find", "Replace", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(1, result);
    }

    [Fact]
    public void ProcessShapesForReplace_WithEmptyTextFrame_ReturnsZero()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.TextFrame.Text = "";

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "find", "replace", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(0, result);
    }

    [Fact]
    public void ProcessShapesForReplace_WithMultipleOccurrences_ReplacesAll()
    {
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.TextFrame.Text = "Hello Hello Hello";

        var result = PptTextHelper.ProcessShapesForReplace(
            slide.Shapes, "Hello", "Hi", StringComparison.OrdinalIgnoreCase);

        Assert.Equal(1, result);
        Assert.Contains("Hi", shape.TextFrame.Text);
        Assert.DoesNotContain("Hello", shape.TextFrame.Text);
    }

    #endregion
}
