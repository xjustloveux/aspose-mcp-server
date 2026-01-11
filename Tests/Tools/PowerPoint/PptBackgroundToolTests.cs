using System.Drawing;
using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptBackgroundTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptBackgroundToolTests : PptTestBase
{
    private readonly PptBackgroundTool _tool;

    public PptBackgroundToolTests()
    {
        _tool = new PptBackgroundTool(SessionManager);
    }

    private string CreatePresentationWithSolidBackground(string fileName, Color bgColor)
    {
        var filePath = CreateTestFilePath(fileName);
        using var ppt = new Presentation();
        ppt.Slides[0].Background.Type = BackgroundType.OwnBackground;
        ppt.Slides[0].Background.FillFormat.FillType = FillType.Solid;
        ppt.Slides[0].Background.FillFormat.SolidFillColor.Color = bgColor;
        ppt.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Set_WithColor_ShouldSetSolidBackground()
    {
        var pptPath = CreatePresentation("test_set_color.pptx");
        var outputPath = CreateTestFilePath("test_set_color_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, color: "#FF0000", outputPath: outputPath);
        Assert.StartsWith("Background", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(FillType.Solid, presentation.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public void Get_ShouldReturnBackgroundInfo()
    {
        var pptPath = CreatePresentation("test_get.pptx");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("hasBackground", out _));
        Assert.True(json.RootElement.TryGetProperty("fillType", out _));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("SET")]
    [InlineData("Set")]
    [InlineData("set")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_set_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_set_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, color: "#FF0000", outputPath: outputPath);
        Assert.StartsWith("Background", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Set_WithSessionId_ShouldSetInMemory()
    {
        var pptPath = CreatePresentation("test_session_set.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, color: "#FF0000");
        Assert.StartsWith("Background", result);
        Assert.Contains("session", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(FillType.Solid, ppt.Slides[0].Background.FillFormat.FillType);
    }

    [Fact]
    public void Get_WithSessionId_ShouldReturnBackgroundInfo()
    {
        var pptPath = CreatePresentation("test_session_get.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("slideIndex", out _));
        Assert.True(json.RootElement.TryGetProperty("fillType", out _));
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session", slideIndex: 0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithSolidBackground("test_path_bg.pptx", Color.Red);
        var pptPath2 = CreatePresentationWithSolidBackground("test_session_bg.pptx", Color.Blue);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId, slideIndex: 0);
        var json = JsonDocument.Parse(result);
        var colorHex = json.RootElement.GetProperty("color").GetString();
        Assert.Contains("0000FF", colorHex!);
    }

    #endregion
}
