using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptTextFormatTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptTextFormatToolTests : PptTestBase
{
    private readonly PptTextFormatTool _tool;

    public PptTextFormatToolTests()
    {
        _tool = new PptTextFormatTool(SessionManager);
    }

    private string CreatePresentationWithTextBox(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        textBox.TextFrame.Text = "Sample Text";
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region Operation Routing

    [Fact]
    public void Execute_WithNoFormattingOptions_ShouldStillSucceed()
    {
        var pptPath = CreatePresentationWithTextBox("test_format_none.pptx");
        var outputPath = CreateTestFilePath("test_format_none_output.pptx");
        var result = _tool.Execute(path: pptPath, outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region File I/O Smoke Tests

    [Fact]
    public void Execute_WithFontOptions_ShouldApplyFontFormatting()
    {
        var pptPath = CreatePresentationWithTextBox("test_format_font.pptx");
        var outputPath = CreateTestFilePath("test_format_font_output.pptx");
        var result = _tool.Execute(path: pptPath, fontName: "Arial", fontSize: 16, bold: true, outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithHexColor_ShouldApplyColor()
    {
        var pptPath = CreatePresentationWithTextBox("test_format_color.pptx");
        var outputPath = CreateTestFilePath("test_format_color_output.pptx");
        var result = _tool.Execute(path: pptPath, color: "#FF0000", outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithAllFormattingOptions_ShouldApplyAllFormats()
    {
        var pptPath = CreatePresentationWithTextBox("test_format_all.pptx");
        var outputPath = CreateTestFilePath("test_format_all_output.pptx");
        var result = _tool.Execute(path: pptPath, fontName: "Arial", fontSize: 14, bold: true, italic: true,
            color: "#0000FF",
            outputPath: outputPath);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Session Management

    [Fact]
    public void Execute_WithSessionId_ShouldFormatInMemory()
    {
        var pptPath = CreatePresentationWithTextBox("test_session_format.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute(sessionId: sessionId, fontName: "Arial", fontSize: 16, bold: true);
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithSessionId_ShouldApplyColorInMemory()
    {
        var pptPath = CreatePresentationWithTextBox("test_session_color.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute(sessionId: sessionId, color: "#FF0000");
        Assert.StartsWith("Batch formatted text applied to", result);
        Assert.Contains("session", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute(sessionId: "invalid_session_id", fontName: "Arial"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestFilePath("test_path_format.pptx");
        using (var pres1 = new Presentation())
        {
            var shape1 = pres1.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape1.TextFrame.Text = "PathText";
            pres1.Save(pptPath1, SaveFormat.Pptx);
        }

        var pptPath2 = CreateTestFilePath("test_session_format2.pptx");
        using (var pres2 = new Presentation())
        {
            var shape2 = pres2.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shape2.TextFrame.Text = "SessionText";
            pres2.Save(pptPath2, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute(path: pptPath1, sessionId: sessionId, fontName: "Arial", bold: true);
        Assert.Contains("session", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.NotNull(ppt);
    }

    #endregion
}
