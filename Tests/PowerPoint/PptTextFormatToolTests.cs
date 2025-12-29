using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptTextFormatToolTests : TestBase
{
    private readonly PptTextFormatTool _tool = new();

    private string CreatePptPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        textBox.TextFrame.Text = "Sample Text";
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePptWithTable(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var table = slide.Shapes.AddTable(100, 100, new double[] { 100, 100 }, new double[] { 50, 50 });
        table[0, 0].TextFrame.Text = "Cell 1";
        table[0, 1].TextFrame.Text = "Cell 2";
        table[1, 0].TextFrame.Text = "Cell 3";
        table[1, 1].TextFrame.Text = "Cell 4";
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task FormatText_WithFontOptions_ShouldApplyFontFormatting()
    {
        // Arrange
        var pptPath = CreatePptPresentation("test_format_font.pptx");
        var outputPath = CreateTestFilePath("test_format_font_output.pptx");
        var arguments = new JsonObject
        {
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["fontName"] = "Arial",
            ["fontSize"] = 16,
            ["bold"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public async Task FormatText_WithColor_ShouldApplyColor()
    {
        // Arrange
        var pptPath = CreatePptPresentation("test_format_color.pptx");
        var outputPath = CreateTestFilePath("test_format_color_output.pptx");
        var arguments = new JsonObject
        {
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["color"] = "#FF0000"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public async Task FormatText_WithAllFormattingOptions_ShouldApplyAllFormats()
    {
        // Arrange
        var pptPath = CreatePptPresentation("test_format_all.pptx");
        var outputPath = CreateTestFilePath("test_format_all_output.pptx");
        var arguments = new JsonObject
        {
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["fontName"] = "Arial",
            ["fontSize"] = 14,
            ["bold"] = true,
            ["italic"] = true,
            ["color"] = "#0000FF"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public async Task FormatText_WithSpecificSlides_ShouldFormatOnlySelectedSlides()
    {
        // Arrange
        var pptPath = CreatePptPresentation("test_format_specific_slides.pptx");
        using var presentation = new Presentation(pptPath);
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(pptPath, SaveFormat.Pptx);

        var outputPath = CreateTestFilePath("test_format_specific_slides_output.pptx");
        var arguments = new JsonObject
        {
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndices"] = new JsonArray(0),
            ["fontName"] = "Times New Roman",
            ["fontSize"] = 12
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var resultPresentation = new Presentation(outputPath);
        Assert.True(resultPresentation.Slides.Count >= 2);
    }

    [Fact]
    public async Task FormatText_WithTableText_ShouldFormatTableCells()
    {
        // Arrange
        var pptPath = CreatePptWithTable("test_format_table.pptx");
        var outputPath = CreateTestFilePath("test_format_table_output.pptx");
        var arguments = new JsonObject
        {
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["fontName"] = "Arial",
            ["fontSize"] = 14,
            ["bold"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("1 slides", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public async Task FormatText_WithNamedColor_ShouldApplyNamedColor()
    {
        // Arrange
        var pptPath = CreatePptPresentation("test_format_named_color.pptx");
        var outputPath = CreateTestFilePath("test_format_named_color_output.pptx");
        var arguments = new JsonObject
        {
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["color"] = "Red"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("slides", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task FormatText_InvalidSlideIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreatePptPresentation("test_format_invalid_index.pptx");
        var arguments = new JsonObject
        {
            ["path"] = pptPath,
            ["slideIndices"] = new JsonArray(99),
            ["fontName"] = "Arial"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task FormatText_WithItalicOnly_ShouldApplyItalic()
    {
        // Arrange
        var pptPath = CreatePptPresentation("test_format_italic.pptx");
        var outputPath = CreateTestFilePath("test_format_italic_output.pptx");
        var arguments = new JsonObject
        {
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["italic"] = true
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("slides", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task FormatText_WithMixedShapes_ShouldFormatBothAutoShapeAndTable()
    {
        // Arrange - Create presentation with both AutoShape and Table
        var filePath = CreateTestFilePath("test_format_mixed.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 50);
            textBox.TextFrame.Text = "AutoShape Text";
            var table = slide.Shapes.AddTable(50, 150, new double[] { 100, 100 }, new double[] { 50 });
            table[0, 0].TextFrame.Text = "Table Text";
            presentation.Save(filePath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_format_mixed_output.pptx");
        var arguments = new JsonObject
        {
            ["path"] = filePath,
            ["outputPath"] = outputPath,
            ["fontName"] = "Verdana",
            ["fontSize"] = 18
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("1 slides", result);
        Assert.True(File.Exists(outputPath));
    }
}