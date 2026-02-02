extern alias SysDrawing;
using System.Drawing;
using System.Runtime.Versioning;
using Aspose.Words;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Word;
using Bitmap = SysDrawing::System.Drawing.Bitmap;
using Brushes = SysDrawing::System.Drawing.Brushes;
using Font = SysDrawing::System.Drawing.Font;
using Graphics = SysDrawing::System.Drawing.Graphics;
using ImageFormat = SysDrawing::System.Drawing.Imaging.ImageFormat;

namespace AsposeMcpServer.Tests.Tools.Word;

/// <summary>
///     Integration tests for WordWatermarkTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
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

    #region File I/O Smoke Tests

    [Fact]
    public void AddTextWatermark_ShouldAddWatermarkWithDefaultOptions()
    {
        var docPath = CreateWordDocument("test_add_watermark.docx");
        var outputPath = CreateTestFilePath("test_add_watermark_output.docx");
        var result = _tool.Execute("add", docPath, outputPath: outputPath,
            text: "CONFIDENTIAL", fontSize: 72, isSemitransparent: true);
        Assert.True(File.Exists(outputPath));
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        var doc = new Document(outputPath);
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
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
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
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        var doc = new Document(outputPath);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("AdD")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var docPath = CreateWordDocument($"test_{operation}_case.docx");
        var outputPath = CreateTestFilePath($"test_{operation}_case_output.docx");
        var result = _tool.Execute(operation, docPath, outputPath: outputPath, text: "TEST");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
        var doc = new Document(outputPath);
        Assert.NotEqual(WatermarkType.None, doc.Watermark.Type);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_invalid_op.docx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", docPath, text: "TEST"));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddTextWatermark_WithSessionId_ShouldAddWatermarkInMemory()
    {
        var docPath = CreateWordDocument("test_session_add_watermark.docx");
        var sessionId = OpenSession(docPath);
        var result = _tool.Execute("add", sessionId: sessionId, text: "SESSION WATERMARK");
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
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
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
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
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        var doc = SessionManager.GetDocument<Document>(sessionId);
        Assert.Equal(WatermarkType.None, doc.Watermark.Type);
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
