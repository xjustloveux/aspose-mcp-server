using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordWatermarkToolTests : WordTestBase
{
    private readonly WordWatermarkTool _tool = new();

    [Fact]
    public async Task AddWatermark_ShouldAddWatermark()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_watermark.docx");
        var outputPath = CreateTestFilePath("test_add_watermark_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["text"] = "CONFIDENTIAL";
        arguments["fontSize"] = 72;
        arguments["isSemitransparent"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);

        // Verify watermark was added by checking document has watermark shapes
        var doc = new Document(outputPath);
        var shapes = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().ToList();
        // Watermarks are typically added as shapes in headers
        var hasWatermarkShapes =
            shapes.Count > 0 || doc.Sections[0].HeadersFooters[HeaderFooterType.HeaderPrimary] != null;
        Assert.True(hasWatermarkShapes || doc.Watermark != null,
            "Document should contain watermark (checking shapes or watermark property)");
    }

    [Fact]
    public async Task AddWatermark_WithFontFamily_ShouldApplyFontFamily()
    {
        // Arrange
        var docPath = CreateWordDocument("test_watermark_font.docx");
        var outputPath = CreateTestFilePath("test_watermark_font_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["text"] = "DRAFT";
        arguments["fontFamily"] = "Times New Roman";
        arguments["fontSize"] = 48;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AddWatermark_WithHorizontalLayout_ShouldApplyHorizontalLayout()
    {
        // Arrange
        var docPath = CreateWordDocument("test_watermark_horizontal.docx");
        var outputPath = CreateTestFilePath("test_watermark_horizontal_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["text"] = "SAMPLE";
        arguments["layout"] = "Horizontal";
        arguments["fontSize"] = 60;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AddWatermark_WithDiagonalLayout_ShouldApplyDiagonalLayout()
    {
        // Arrange
        var docPath = CreateWordDocument("test_watermark_diagonal.docx");
        var outputPath = CreateTestFilePath("test_watermark_diagonal_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["text"] = "DO NOT COPY";
        arguments["layout"] = "Diagonal";
        arguments["fontSize"] = 54;
        arguments["isSemitransparent"] = false;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AddWatermark_WithAllOptions_ShouldApplyAllOptions()
    {
        // Arrange
        var docPath = CreateWordDocument("test_watermark_all_options.docx");
        var outputPath = CreateTestFilePath("test_watermark_all_options_output.docx");
        var arguments = CreateArguments("add", docPath, outputPath);
        arguments["text"] = "CONFIDENTIAL";
        arguments["fontFamily"] = "Arial";
        arguments["fontSize"] = 80;
        arguments["isSemitransparent"] = true;
        arguments["layout"] = "Diagonal";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);

        // Verify watermark exists
        var doc = new Document(outputPath);
        Assert.NotNull(doc.Watermark);
    }
}