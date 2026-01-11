using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptLayoutTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptLayoutToolTests : PptTestBase
{
    private readonly PptLayoutTool _tool;

    public PptLayoutToolTests()
    {
        _tool = new PptLayoutTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void GetMasters_ShouldReturnMasterSlides()
    {
        var pptPath = CreatePresentation("test_get_masters.pptx");
        var result = _tool.Execute("get_masters", pptPath);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
        Assert.Contains("layoutType", result);
    }

    [Fact]
    public void GetLayouts_ShouldReturnLayoutsWithType()
    {
        var pptPath = CreatePresentation("test_get_layouts.pptx");
        var result = _tool.Execute("get_layouts", pptPath);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.TryGetProperty("mastersCount", out _));
        Assert.Contains("layoutType", result);
    }

    [Fact]
    public void Set_ShouldSetSlideLayout()
    {
        var pptPath = CreatePresentation("test_set_layout.pptx");
        var outputPath = CreateTestFilePath("test_set_layout_output.pptx");
        var result = _tool.Execute("set", pptPath, slideIndex: 0, layout: "Blank", outputPath: outputPath);
        Assert.StartsWith("Layout 'Blank' set for slide 0", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideLayoutType.Blank, presentation.Slides[0].LayoutSlide.LayoutType);
    }

    [Fact]
    public void ApplyMaster_ShouldApplyToAllSlides()
    {
        var pptPath = CreatePresentation("test_apply_master.pptx", 3);
        var outputPath = CreateTestFilePath("test_apply_master_output.pptx");
        var result = _tool.Execute("apply_master", pptPath, masterIndex: 0, layoutIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Master", result);
        Assert.Contains("3 slides", result);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_MASTERS")]
    [InlineData("Get_Masters")]
    [InlineData("get_masters")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_masters_{operation.Replace("_", "")}.pptx");
        var result = _tool.Execute(operation, pptPath);
        Assert.Contains("\"count\"", result);
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
    public void GetMasters_WithSessionId_ShouldReturnMasterSlides()
    {
        var pptPath = CreatePresentation("test_session_get_masters.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_masters", sessionId: sessionId);
        var json = JsonDocument.Parse(result);
        Assert.True(json.RootElement.GetProperty("count").GetInt32() > 0);
        Assert.Contains("layoutType", result);
    }

    [Fact]
    public void Set_WithSessionId_ShouldSetLayoutInMemory()
    {
        var pptPath = CreatePresentation("test_session_set_layout.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, layout: "Blank");
        Assert.StartsWith("Layout 'Blank' set for slide 0", result);
        Assert.Contains("session", result);
        Assert.Equal(SlideLayoutType.Blank, ppt.Slides[0].LayoutSlide.LayoutType);
    }

    [Fact]
    public void ApplyMaster_WithSessionId_ShouldApplyInMemory()
    {
        var pptPath = CreatePresentation("test_session_apply_master.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("apply_master", sessionId: sessionId, masterIndex: 0, layoutIndex: 0);
        Assert.StartsWith("Master", result);
        Assert.Contains("3 slides", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_masters", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentation("test_path_layout.pptx");
        var pptPath2 = CreatePresentation("test_session_layout.pptx", 5);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("apply_master", pptPath1, sessionId, masterIndex: 0, layoutIndex: 0);
        Assert.StartsWith("Master", result);
        Assert.Contains("5 slides", result);
        Assert.Contains("session", result);
    }

    #endregion
}
