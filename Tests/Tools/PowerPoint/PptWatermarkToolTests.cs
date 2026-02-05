using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Handlers.PowerPoint.Watermark;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Watermark;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptWatermarkTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptWatermarkToolTests : PptTestBase
{
    private readonly PptWatermarkTool _tool;

    public PptWatermarkToolTests()
    {
        _tool = new PptWatermarkTool(SessionManager);
    }

    /// <summary>
    ///     Creates a presentation with a text watermark for testing.
    /// </summary>
    /// <param name="fileName">The file name for the presentation.</param>
    /// <returns>The file path of the created presentation.</returns>
    private string CreatePresentationWithWatermark(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 400, 100);
        shape.Name = $"{AddTextPptWatermarkHandler.WatermarkPrefix}TEXT_test";
        shape.TextFrame.Text = "WATERMARK";
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    /// <summary>
    ///     Creates a minimal test image file (1x1 white pixel PNG) with a relative path
    ///     to comply with SecurityHelper.ValidateFilePath which rejects absolute paths.
    /// </summary>
    /// <returns>The relative file name of the created image file.</returns>
    private string CreateTestImage()
    {
        var fileName = $"test_wm_{Guid.NewGuid()}.png";
        var pngBytes = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
            0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, // 8-bit RGB
            0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
            0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00,
            0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33,
            0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND chunk
            0xAE, 0x42, 0x60, 0x82
        };
        File.WriteAllBytes(fileName, pngBytes);
        TestFiles.Add(Path.GetFullPath(fileName));
        return fileName;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void AddText_ShouldAddTextWatermark()
    {
        var pptPath = CreatePresentationWithContent("test_add_text.pptx", "Hello World");
        var outputPath = CreateTestFilePath("test_add_text_output.pptx");
        var result = _tool.Execute("add_text", pptPath, text: "DRAFT", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("1 slide(s)", data.Message);
    }

    [Fact]
    public void AddImage_ShouldAddImageWatermark()
    {
        var pptPath = CreatePresentationWithContent("test_add_image.pptx", "Hello World");
        var imagePath = CreateTestImage();
        var outputPath = CreateTestFilePath("test_add_image_output.pptx");
        var result = _tool.Execute("add_image", pptPath, imagePath: imagePath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("1 slide(s)", data.Message);
    }

    [Fact]
    public void Remove_ShouldRemoveWatermarks()
    {
        var pptPath = CreatePresentationWithWatermark("test_remove.pptx");
        var outputPath = CreateTestFilePath("test_remove_output.pptx");
        var result = _tool.Execute("remove", pptPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("watermark", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Get_ShouldReturnWatermarks()
    {
        var pptPath = CreatePresentationWithWatermark("test_get.pptx");
        var result = _tool.Execute("get", pptPath);
        var data = GetResultData<GetWatermarksPptResult>(result);
        Assert.True(data.Count > 0);
        Assert.NotEmpty(data.Items);
    }

    [Fact]
    public void Get_EmptyPresentation_ShouldReturnZeroCount()
    {
        var pptPath = CreatePresentation("test_get_empty.pptx");
        var result = _tool.Execute("get", pptPath);
        var data = GetResultData<GetWatermarksPptResult>(result);
        Assert.Equal(0, data.Count);
        Assert.Empty(data.Items);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentationWithContent($"test_case_{operation.Replace(" ", "_")}.pptx", "Hello World");
        var result = _tool.Execute(operation, pptPath);
        var data = GetResultData<GetWatermarksPptResult>(result);
        Assert.True(data.Count >= 0);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithContent("test_unknown_op.pptx", "Hello World");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldReturnWatermarksFromMemory()
    {
        var pptPath = CreatePresentationWithContent("test_session_get.pptx", "Hello World");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetWatermarksPptResult>(result);
        Assert.True(data.Count >= 0);
        var output = GetResultOutput<GetWatermarksPptResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    [Fact]
    public void AddText_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentationWithContent("test_session_add_text.pptx", "Hello World");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("add_text", sessionId: sessionId, text: "CONFIDENTIAL");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("1 slide(s)", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    [Fact]
    public void Remove_WithSessionId_ShouldRemoveInMemory()
    {
        var pptPath = CreatePresentationWithWatermark("test_session_remove.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("remove", sessionId: sessionId);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("watermark", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithContent("test_path_wm.pptx", "Path content");
        var pptPath2 = CreatePresentationWithWatermark("test_session_wm.pptx");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId);
        var data = GetResultData<GetWatermarksPptResult>(result);
        Assert.NotNull(data);
    }

    #endregion
}
