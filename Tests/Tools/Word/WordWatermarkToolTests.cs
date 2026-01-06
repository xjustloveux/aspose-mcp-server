using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.Versioning;
using Aspose.Words;
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

    #region General

    [Fact]
    public void AddTextWatermark_ShouldAddWatermarkWithDefaultOptions()
    {
        var docPath = CreateWordDocument("test_add_watermark.docx");
        var outputPath = CreateTestFilePath("test_add_watermark_output.docx");

        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "CONFIDENTIAL", fontSize: 72, isSemitransparent: true);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Text watermark added to document", result);

        var doc = new Document(outputPath);
        Assert.NotEqual(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void AddTextWatermark_WithFontFamily_ShouldApplyFontFamily()
    {
        var docPath = CreateWordDocument("test_watermark_font.docx");
        var outputPath = CreateTestFilePath("test_watermark_font_output.docx");

        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "DRAFT", fontFamily: "Times New Roman", fontSize: 48);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Text watermark added to document", result);
    }

    [Theory]
    [InlineData("Horizontal")]
    [InlineData("Diagonal")]
    public void AddTextWatermark_WithLayout_ShouldApplyLayout(string layout)
    {
        var docPath = CreateWordDocument($"test_watermark_{layout.ToLower()}.docx");
        var outputPath = CreateTestFilePath($"test_watermark_{layout.ToLower()}_output.docx");

        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "SAMPLE", layout: layout, fontSize: 60);

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Text watermark added to document", result);

        var doc = new Document(outputPath);
        Assert.NotEqual(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void AddTextWatermark_WithAllOptions_ShouldApplyAllOptions()
    {
        var docPath = CreateWordDocument("test_watermark_all_options.docx");
        var outputPath = CreateTestFilePath("test_watermark_all_options_output.docx");

        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "CONFIDENTIAL", fontFamily: "Arial", fontSize: 80,
            isSemitransparent: true, layout: "Diagonal");

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Text watermark added to document", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Text, doc.Watermark.Type);
    }

    [Fact]
    public void AddTextWatermark_WithoutOutputPath_ShouldOverwriteInput()
    {
        var docPath = CreateWordDocument("test_watermark_overwrite.docx");

        var result = _tool.Execute("add", docPath, text: "OVERWRITE TEST");

        Assert.StartsWith("Text watermark added to document", result);
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

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Image watermark added to document", result);
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

        Assert.True(File.Exists(outputPath));
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

        Assert.True(File.Exists(outputPath));
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

        Assert.True(File.Exists(outputPath));
        Assert.StartsWith("Image watermark added to document", result);
        Assert.Contains("Scale: 0.75", result);
        Assert.Contains("Washout: True", result);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public void RemoveWatermark_ShouldRemoveTextWatermark()
    {
        var docPath = CreateWordDocument("test_remove_watermark.docx");
        var watermarkedPath = CreateTestFilePath("test_remove_watermark_with.docx");
        var outputPath = CreateTestFilePath("test_remove_watermark_output.docx");

        _tool.Execute("add", docPath, outputPath: watermarkedPath, text: "TO BE REMOVED");

        var result = _tool.Execute("remove", watermarkedPath, outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void RemoveWatermark_ShouldRemoveImageWatermark()
    {
        var docPath = CreateWordDocument("test_remove_image_watermark.docx");
        var imagePath = CreateTestImage("watermark_to_remove.png");
        var watermarkedPath = CreateTestFilePath("test_remove_image_watermark_with.docx");
        var outputPath = CreateTestFilePath("test_remove_image_watermark_output.docx");

        _tool.Execute("add_image", docPath, outputPath: watermarkedPath, imagePath: imagePath);

        var docWithWatermark = new Document(watermarkedPath);
        Assert.Equal(WatermarkType.Image, docWithWatermark.Watermark.Type);

        var result = _tool.Execute("remove", watermarkedPath, outputPath: outputPath);

        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);

        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void RemoveWatermark_WhenNoWatermarkExists_ShouldReturnNotFoundMessage()
    {
        var docPath = CreateWordDocument("test_remove_no_watermark.docx");
        var outputPath = CreateTestFilePath("test_remove_no_watermark_output.docx");

        var result = _tool.Execute("remove", docPath, outputPath: outputPath);

        Assert.Contains("No watermark found", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("AdD")]
    [InlineData("add")]
    public void Execute_OperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation}_case.docx");
        var outputPath = CreateTestFilePath($"test_{operation}_case_output.docx");

        var result = _tool.Execute(operation, docPath, outputPath: outputPath, text: "TEST");

        Assert.StartsWith("Text watermark added to document", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("ADD_IMAGE")]
    [InlineData("Add_Image")]
    [InlineData("add_image")]
    public void Execute_AddImageOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation.Replace("_", "")}_case.docx");
        var imagePath = CreateTestImage($"watermark_{operation.Replace("_", "")}.png");
        var outputPath = CreateTestFilePath($"test_{operation.Replace("_", "")}_case_output.docx");

        var result = _tool.Execute(operation, docPath, outputPath: outputPath, imagePath: imagePath);

        Assert.StartsWith("Image watermark added to document", result);
        Assert.True(File.Exists(outputPath));
    }

    [Theory]
    [InlineData("REMOVE")]
    [InlineData("ReMoVe")]
    [InlineData("remove")]
    public void Execute_RemoveOperationIsCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation}_remove_case.docx");
        var tempPath = CreateTestFilePath($"test_{operation}_remove_case_temp.docx");
        _tool.Execute("add", docPath, outputPath: tempPath, text: "TO REMOVE");

        var outputPath = CreateTestFilePath($"test_{operation}_remove_case_output.docx");

        var result = _tool.Execute(operation, tempPath, outputPath: outputPath);

        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("HORIZONTAL")]
    [InlineData("DiAgOnAl")]
    [InlineData("horizontal")]
    public void AddTextWatermark_LayoutIsCaseInsensitive(string layout)
    {
        var docPath = CreateWordDocument($"test_layout_{layout.ToLower()}.docx");
        var outputPath = CreateTestFilePath($"test_layout_{layout.ToLower()}_output.docx");

        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "LAYOUT TEST", layout: layout);

        Assert.StartsWith("Text watermark added to document", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Exception

    [Theory]
    [InlineData("unknown_operation")]
    [InlineData("invalid")]
    [InlineData("")]
    public void Execute_WithInvalidOperation_ShouldThrowArgumentException(string operation)
    {
        var docPath = CreateWordDocument($"test_invalid_op_{operation.GetHashCode()}.docx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute(operation, docPath, text: "TEST"));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void AddTextWatermark_WithNullOrEmptyText_ShouldThrowArgumentException(string? text)
    {
        var docPath = CreateWordDocument($"test_invalid_text_{text?.GetHashCode() ?? 0}.docx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", docPath, text: text));

        Assert.Contains("Text is required", ex.Message);
    }

    [Fact]
    public void AddTextWatermark_WithWhitespaceOnlyText_ShouldThrowException()
    {
        var docPath = CreateWordDocument("test_whitespace_text.docx");

        Assert.ThrowsAny<ArgumentException>(() =>
            _tool.Execute("add", docPath, text: "   "));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void AddImageWatermark_WithInvalidImagePath_ShouldThrowArgumentException(string? imagePath)
    {
        var docPath = CreateWordDocument($"test_invalid_image_path_{imagePath?.GetHashCode() ?? 0}.docx");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_image", docPath, imagePath: imagePath));

        Assert.Contains("imagePath is required", ex.Message);
    }

    [Fact]
    public void AddImageWatermark_WithNonExistentFile_ShouldThrowFileNotFoundException()
    {
        var docPath = CreateWordDocument("test_nonexistent_image.docx");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("add_image", docPath, imagePath: "nonexistent_image.png"));
    }

    [Fact]
    public void AddImageWatermark_WithNonImageFile_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_non_image.docx");
        var textFilePath = CreateTestFilePath("not_an_image.txt");
        File.WriteAllText(textFilePath, "This is not an image");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add_image", docPath, imagePath: textFilePath));

        Assert.Contains("decode", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNeitherPathNorSessionId_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", text: "TEST"));

        Assert.Contains("path", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session

    [Fact]
    public void AddTextWatermark_WithSessionId_ShouldAddWatermarkInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_watermark.docx");
        var sessionId = OpenSession(docPath);

        var result = _tool.Execute("add", sessionId: sessionId, text: "SESSION WATERMARK");

        Assert.StartsWith("Text watermark added to document", result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotEqual(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void AddTextWatermark_WithSessionId_AndLayoutOptions_ShouldWork()
    {
        var docPath = CreateWordDocument("test_session_horizontal.docx");
        var sessionId = OpenSession(docPath);

        var result = _tool.Execute("add", sessionId: sessionId, text: "HORIZONTAL SESSION",
            layout: "Horizontal", fontSize: 60);

        Assert.StartsWith("Text watermark added to document", result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotEqual(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void AddImageWatermark_WithSessionId_ShouldAddImageWatermarkInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_image_watermark.docx");
        var imagePath = CreateTestImage("session_watermark.png");
        var sessionId = OpenSession(docPath);

        var result = _tool.Execute("add_image", sessionId: sessionId, imagePath: imagePath);

        Assert.StartsWith("Image watermark added to document", result);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(WatermarkType.Image, doc.Watermark.Type);
    }

    [Fact]
    public void RemoveWatermark_WithSessionId_ShouldRemoveWatermarkInMemory()
    {
        var docPath = CreateWordDocument("test_session_remove_watermark.docx");
        var tempPath = CreateTestFilePath("test_session_remove_watermark_temp.docx");
        _tool.Execute("add", docPath, outputPath: tempPath, text: "TO BE REMOVED");

        var sessionId = OpenSession(tempPath);

        var result = _tool.Execute("remove", sessionId: sessionId);

        Assert.Contains("removed", result, StringComparison.OrdinalIgnoreCase);

        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void RemoveWatermark_WithSessionId_WhenNoWatermark_ShouldReturnMessage()
    {
        var docPath = CreateWordDocument("test_session_remove_none.docx");
        var sessionId = OpenSession(docPath);

        var result = _tool.Execute("remove", sessionId: sessionId);

        Assert.Contains("No watermark found", result);
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

        _tool.Execute("add", docPath1, text: "PATH WATERMARK");

        var sessionId = OpenSession(docPath2);

        _tool.Execute("add", docPath1, sessionId, text: "SESSION WATERMARK");

        var sessionDoc = SessionManager.GetDocument<Document>(sessionId);
        Assert.NotEqual(WatermarkType.None, sessionDoc.Watermark.Type);

        var diskDoc2 = new Document(docPath2);
        Assert.Equal(WatermarkType.None, diskDoc2.Watermark.Type);
    }

    #endregion
}