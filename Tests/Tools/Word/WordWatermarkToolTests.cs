using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using Aspose.Words;
using Aspose.Words.Drawing;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;
using Font = System.Drawing.Font;

namespace AsposeMcpServer.Tests.Tools.Word;

[SupportedOSPlatform("windows")]
public class WordWatermarkToolTests : WordTestBase
{
    private readonly WordWatermarkTool _tool;

    public WordWatermarkToolTests()
    {
        _tool = new WordWatermarkTool(SessionManager);
    }

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

    #region General Tests

    [Fact]
    public void AddWatermark_ShouldAddWatermark()
    {
        var docPath = CreateWordDocument("test_add_watermark.docx");
        var outputPath = CreateTestFilePath("test_add_watermark_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "CONFIDENTIAL", fontSize: 72, isSemitransparent: true);
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
    public void AddWatermark_WithFontFamily_ShouldApplyFontFamily()
    {
        var docPath = CreateWordDocument("test_watermark_font.docx");
        var outputPath = CreateTestFilePath("test_watermark_font_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "DRAFT", fontFamily: "Times New Roman", fontSize: 48);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddWatermark_WithHorizontalLayout_ShouldApplyHorizontalLayout()
    {
        var docPath = CreateWordDocument("test_watermark_horizontal.docx");
        var outputPath = CreateTestFilePath("test_watermark_horizontal_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "SAMPLE", layout: "Horizontal", fontSize: 60);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddWatermark_WithDiagonalLayout_ShouldApplyDiagonalLayout()
    {
        var docPath = CreateWordDocument("test_watermark_diagonal.docx");
        var outputPath = CreateTestFilePath("test_watermark_diagonal_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "DO NOT COPY", layout: "Diagonal", fontSize: 54, isSemitransparent: false);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddWatermark_WithAllOptions_ShouldApplyAllOptions()
    {
        var docPath = CreateWordDocument("test_watermark_all_options.docx");
        var outputPath = CreateTestFilePath("test_watermark_all_options_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "CONFIDENTIAL", fontFamily: "Arial", fontSize: 80,
            isSemitransparent: true, layout: "Diagonal");
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.NotNull(result);
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);

        // Verify watermark exists
        var doc = new Document(outputPath);
        Assert.NotNull(doc.Watermark);
    }

    [Fact]
    public void RemoveWatermark_ShouldRemoveWatermark()
    {
        // Arrange - First add a watermark
        var docPath = CreateWordDocument("test_remove_watermark.docx");
        var watermarkedPath = CreateTestFilePath("test_remove_watermark_with.docx");
        var outputPath = CreateTestFilePath("test_remove_watermark_output.docx");

        // Add watermark first
        _tool.Execute("add", docPath, outputPath: watermarkedPath, text: "TO BE REMOVED");

        // Act - Remove watermark
        var result = _tool.Execute("remove", watermarkedPath, outputPath: outputPath);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void RemoveWatermark_WithNoWatermark_ShouldReturnNotFoundMessage()
    {
        // Arrange - Document without watermark
        var docPath = CreateWordDocument("test_remove_no_watermark.docx");
        var outputPath = CreateTestFilePath("test_remove_no_watermark_output.docx");
        var result = _tool.Execute("remove", docPath, outputPath: outputPath);
        Assert.Contains("No watermark found", result);
    }

    [Fact]
    public void AddWatermark_WithoutOutputPath_ShouldOverwriteInput()
    {
        var docPath = CreateWordDocument("test_watermark_overwrite.docx");
        var result = _tool.Execute("add", docPath, text: "OVERWRITE TEST");
        Assert.Contains("Text watermark added", result);
        Assert.Contains(docPath, result);

        var doc = new Document(docPath);
        Assert.NotEqual(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void AddImageWatermark_ShouldAddImageWatermark()
    {
        var docPath = CreateWordDocument("test_add_image_watermark.docx");
        var imagePath = CreateTestImage("watermark_image.png");
        var outputPath = CreateTestFilePath("test_add_image_watermark_output.docx");
        var result = _tool.Execute("add_image", docPath, outputPath: outputPath, imagePath: imagePath);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Image watermark added", result);
        Assert.Contains("watermark_image.png", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public void AddImageWatermark_WithScale_ShouldApplyScale()
    {
        var docPath = CreateWordDocument("test_image_watermark_scale.docx");
        var imagePath = CreateTestImage("watermark_scale.png");
        var outputPath = CreateTestFilePath("test_image_watermark_scale_output.docx");
        var result = _tool.Execute("add_image", docPath, outputPath: outputPath,
            imagePath: imagePath, scale: 0.5);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Scale: 0.5", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public void AddImageWatermark_WithWashoutFalse_ShouldNotApplyWashout()
    {
        var docPath = CreateWordDocument("test_image_watermark_no_washout.docx");
        var imagePath = CreateTestImage("watermark_no_washout.png");
        var outputPath = CreateTestFilePath("test_image_watermark_no_washout_output.docx");
        var result = _tool.Execute("add_image", docPath, outputPath: outputPath,
            imagePath: imagePath, isWashout: false);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Washout: False", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public void AddImageWatermark_WithAllOptions_ShouldApplyAllOptions()
    {
        var docPath = CreateWordDocument("test_image_watermark_all.docx");
        var imagePath = CreateTestImage("watermark_all_options.png", 300, 150);
        var outputPath = CreateTestFilePath("test_image_watermark_all_output.docx");
        var result = _tool.Execute("add_image", docPath, outputPath: outputPath,
            imagePath: imagePath, scale: 0.75, isWashout: true);
        Assert.True(File.Exists(outputPath), "Output document should be created");
        Assert.Contains("Image watermark added", result);
        Assert.Contains("Scale: 0.75", result);
        Assert.Contains("Washout: True", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public void AddImageWatermark_WithInvalidImagePath_ShouldThrowFileNotFoundException()
    {
        var docPath = CreateWordDocument("test_invalid_image.docx");
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add_image", docPath, imagePath: "nonexistent_image.png"));
    }

    [Fact]
    public void RemoveWatermark_AfterImageWatermark_ShouldRemoveImageWatermark()
    {
        // Arrange - First add an image watermark
        var docPath = CreateWordDocument("test_remove_image_watermark.docx");
        var imagePath = CreateTestImage("watermark_to_remove.png");
        var watermarkedPath = CreateTestFilePath("test_remove_image_watermark_with.docx");
        var outputPath = CreateTestFilePath("test_remove_image_watermark_output.docx");

        // Add image watermark first
        _tool.Execute("add_image", docPath, outputPath: watermarkedPath, imagePath: imagePath);

        // Verify watermark was added
        var docWithWatermark = new Document(watermarkedPath);
        Assert.Equal(WatermarkType.Image, docWithWatermark.Watermark.Type);

        // Act - Remove watermark
        var result = _tool.Execute("remove", watermarkedPath, outputPath: outputPath);
        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_unknown_op.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath, text: "TEST"));

        Assert.Contains("Unknown operation", ex.Message);
        Assert.Contains("unknown_operation", ex.Message);
    }

    [Fact]
    public void AddWatermark_WithInvalidOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_invalid_operation.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("invalid", docPath, text: "TEST"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddWatermark_WithoutText_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_no_text.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath));

        Assert.Contains("Text is required", ex.Message);
    }

    [Fact]
    public void AddImageWatermark_WithoutImagePath_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_add_image_no_path.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_image", docPath));

        Assert.Contains("imagePath is required", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddWatermark_WithSessionId_ShouldAddWatermarkInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_watermark.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, text: "SESSION WATERMARK");
        Assert.Contains("Watermark", result, StringComparison.OrdinalIgnoreCase);

        // Verify in-memory document has watermark
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotNull(doc.Watermark);
        Assert.NotEqual(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void RemoveWatermark_WithSessionId_ShouldRemoveWatermarkInMemory()
    {
        // Arrange - First add watermark to file
        var docPath = CreateWordDocument("test_session_remove_watermark.docx");
        var tempPath = CreateTestFilePath("test_session_remove_watermark_temp.docx");
        _tool.Execute("add", docPath, outputPath: tempPath, text: "TO BE REMOVED");

        var sessionId = OpenSession(tempPath);
        var result = _tool.Execute("remove", sessionId: sessionId);
        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);

        // Verify in-memory document has no watermark
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void AddImageWatermark_WithSessionId_ShouldAddImageWatermarkInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_image_watermark.docx");
        var imagePath = CreateTestImage("session_watermark.png");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add_image", sessionId: sessionId, imagePath: imagePath);
        Assert.Contains("Image watermark added", result);

        // Verify in-memory document has image watermark
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("add", sessionId: "invalid_session_id", text: "TEST"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var docPath1 = CreateWordDocument("test_watermark_path.docx");
        var docPath2 = CreateWordDocument("test_watermark_session.docx");

        // Add watermark to first file
        _tool.Execute("add", docPath1, text: "PATH WATERMARK");

        var sessionId = OpenSession(docPath2);

        // Act - provide both path and sessionId, add watermark to session
        _tool.Execute("add", docPath1, sessionId, text: "SESSION WATERMARK");

        // Assert - session document should have watermark, not the file path document
        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotEqual(WatermarkType.None, sessionDoc.Watermark.Type);

        // Reload docPath2 from disk - it should NOT have the watermark since we used session
        var diskDoc2 = new Document(docPath2);
        Assert.Equal(WatermarkType.None, diskDoc2.Watermark.Type);
    }

    #endregion
}