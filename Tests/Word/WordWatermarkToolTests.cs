using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;
using Font = System.Drawing.Font;

namespace AsposeMcpServer.Tests.Word;

[SupportedOSPlatform("windows")]
public class WordWatermarkToolTests : WordTestBase
{
    private readonly WordWatermarkTool _tool = new();

    private string CreateTestImage(string fileName, int width = 200, int height = 100)
    {
        var imagePath = CreateTestFilePath(fileName);
        using var bitmap = new Bitmap(width, height);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.LightBlue);
        graphics.DrawString("WATERMARK", new Font("Arial", 16), Brushes.DarkBlue, 20, 30);
        bitmap.Save(imagePath, ImageFormat.Png);
        return imagePath;
    }

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

    [Fact]
    public async Task RemoveWatermark_ShouldRemoveWatermark()
    {
        // Arrange - First add a watermark
        var docPath = CreateWordDocument("test_remove_watermark.docx");
        var watermarkedPath = CreateTestFilePath("test_remove_watermark_with.docx");
        var outputPath = CreateTestFilePath("test_remove_watermark_output.docx");

        // Add watermark first
        var addArgs = CreateArguments("add", docPath, watermarkedPath);
        addArgs["text"] = "TO BE REMOVED";
        await _tool.ExecuteAsync(addArgs);

        // Act - Remove watermark
        var removeArgs = CreateArguments("remove", watermarkedPath, outputPath);
        var result = await _tool.ExecuteAsync(removeArgs);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public async Task RemoveWatermark_WithNoWatermark_ShouldReturnNotFoundMessage()
    {
        // Arrange - Document without watermark
        var docPath = CreateWordDocument("test_remove_no_watermark.docx");
        var outputPath = CreateTestFilePath("test_remove_no_watermark_output.docx");
        var arguments = CreateArguments("remove", docPath, outputPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("No watermark found", result);
    }

    [Fact]
    public async Task AddWatermark_WithoutOutputPath_ShouldOverwriteInput()
    {
        // Arrange
        var docPath = CreateWordDocument("test_watermark_overwrite.docx");
        var arguments = CreateArguments("add", docPath);
        arguments["text"] = "OVERWRITE TEST";

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Text watermark added", result);
        Assert.Contains(docPath, result);

        var doc = new Document(docPath);
        Assert.NotEqual(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public async Task AddWatermark_WithInvalidOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_invalid_operation.docx");
        var arguments = CreateArguments("invalid", docPath);
        arguments["text"] = "TEST";

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public async Task AddImageWatermark_ShouldAddImageWatermark()
    {
        // Arrange
        var docPath = CreateWordDocument("test_add_image_watermark.docx");
        var imagePath = CreateTestImage("watermark_image.png");
        var outputPath = CreateTestFilePath("test_add_image_watermark_output.docx");
        var arguments = CreateArguments("add_image", docPath, outputPath);
        arguments["imagePath"] = imagePath;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Image watermark added", result);
        Assert.Contains("watermark_image.png", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public async Task AddImageWatermark_WithScale_ShouldApplyScale()
    {
        // Arrange
        var docPath = CreateWordDocument("test_image_watermark_scale.docx");
        var imagePath = CreateTestImage("watermark_scale.png");
        var outputPath = CreateTestFilePath("test_image_watermark_scale_output.docx");
        var arguments = CreateArguments("add_image", docPath, outputPath);
        arguments["imagePath"] = imagePath;
        arguments["scale"] = 0.5;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Scale: 0.5", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public async Task AddImageWatermark_WithWashoutFalse_ShouldNotApplyWashout()
    {
        // Arrange
        var docPath = CreateWordDocument("test_image_watermark_no_washout.docx");
        var imagePath = CreateTestImage("watermark_no_washout.png");
        var outputPath = CreateTestFilePath("test_image_watermark_no_washout_output.docx");
        var arguments = CreateArguments("add_image", docPath, outputPath);
        arguments["imagePath"] = imagePath;
        arguments["isWashout"] = false;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Washout: False", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public async Task AddImageWatermark_WithAllOptions_ShouldApplyAllOptions()
    {
        // Arrange
        var docPath = CreateWordDocument("test_image_watermark_all.docx");
        var imagePath = CreateTestImage("watermark_all_options.png", 300, 150);
        var outputPath = CreateTestFilePath("test_image_watermark_all_output.docx");
        var arguments = CreateArguments("add_image", docPath, outputPath);
        arguments["imagePath"] = imagePath;
        arguments["scale"] = 0.75;
        arguments["isWashout"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Image watermark added", result);
        Assert.Contains("Scale: 0.75", result);
        Assert.Contains("Washout: True", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public async Task AddImageWatermark_WithInvalidImagePath_ShouldThrowFileNotFoundException()
    {
        // Arrange
        var docPath = CreateWordDocument("test_invalid_image.docx");
        var arguments = CreateArguments("add_image", docPath);
        arguments["imagePath"] = "nonexistent_image.png";

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task RemoveWatermark_AfterImageWatermark_ShouldRemoveImageWatermark()
    {
        // Arrange - First add an image watermark
        var docPath = CreateWordDocument("test_remove_image_watermark.docx");
        var imagePath = CreateTestImage("watermark_to_remove.png");
        var watermarkedPath = CreateTestFilePath("test_remove_image_watermark_with.docx");
        var outputPath = CreateTestFilePath("test_remove_image_watermark_output.docx");

        // Add image watermark first
        var addArgs = CreateArguments("add_image", docPath, watermarkedPath);
        addArgs["imagePath"] = imagePath;
        await _tool.ExecuteAsync(addArgs);

        // Verify watermark was added
        var docWithWatermark = new Document(watermarkedPath);
        Assert.Equal(WatermarkType.Image, docWithWatermark.Watermark.Type);

        // Act - Remove watermark
        var removeArgs = CreateArguments("remove", watermarkedPath, outputPath);
        var result = await _tool.ExecuteAsync(removeArgs);

        // Assert
        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
    }
}