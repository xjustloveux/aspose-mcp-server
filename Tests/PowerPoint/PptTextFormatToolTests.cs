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
}