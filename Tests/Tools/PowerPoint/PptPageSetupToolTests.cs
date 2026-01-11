using Aspose.Slides;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptPageSetupTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptPageSetupToolTests : PptTestBase
{
    private readonly PptPageSetupTool _tool;

    public PptPageSetupToolTests()
    {
        _tool = new PptPageSetupTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void SetSize_WithPreset_ShouldSetSlideSize()
    {
        var pptPath = CreatePresentation("test_set_size.pptx");
        var outputPath = CreateTestFilePath("test_set_size_output.pptx");
        var result = _tool.Execute("set_size", pptPath, preset: "OnScreen16x9", outputPath: outputPath);
        Assert.StartsWith("Slide size set to", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(SlideSizeType.OnScreen, presentation.SlideSize.Type);
    }

    [Fact]
    public void SetOrientation_Portrait_ShouldSwapToPortrait()
    {
        var pptPath = CreatePresentation("test_set_portrait.pptx");
        var outputPath = CreateTestFilePath("test_set_portrait_output.pptx");
        var result = _tool.Execute("set_orientation", pptPath, orientation: "Portrait", outputPath: outputPath);
        Assert.StartsWith("Slide orientation set to", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.SlideSize.Size.Height > presentation.SlideSize.Size.Width);
    }

    [Fact]
    public void SetFooter_ShouldSetFooterText()
    {
        var pptPath = CreatePresentation("test_set_footer.pptx");
        var outputPath = CreateTestFilePath("test_set_footer_output.pptx");
        var result = _tool.Execute("set_footer", pptPath, footerText: "Footer Text", outputPath: outputPath);
        Assert.StartsWith("Footer settings updated for", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void SetSlideNumbering_Show_ShouldShowNumbers()
    {
        var pptPath = CreatePresentation("test_numbering_show.pptx");
        var outputPath = CreateTestFilePath("test_numbering_show_output.pptx");
        var result = _tool.Execute("set_slide_numbering", pptPath, showSlideNumber: true, outputPath: outputPath);
        Assert.StartsWith("Slide numbers", result);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("SET_SIZE")]
    [InlineData("Set_Size")]
    [InlineData("set_size")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_size_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_size_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, preset: "OnScreen16x9", outputPath: outputPath);
        Assert.StartsWith("Slide size set to", result);
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
    public void SetSize_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreatePresentation("test_session_size.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_size", sessionId: sessionId, preset: "OnScreen16x9");
        Assert.StartsWith("Slide size set to", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(SlideSizeType.OnScreen, ppt.SlideSize.Type);
    }

    [Fact]
    public void SetOrientation_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreatePresentation("test_session_orientation.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_orientation", sessionId: sessionId, orientation: "Portrait");
        Assert.StartsWith("Slide orientation set to", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.True(ppt.SlideSize.Size.Height > ppt.SlideSize.Size.Width);
    }

    [Fact]
    public void SetFooter_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreatePresentation("test_session_footer.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_footer", sessionId: sessionId, footerText: "Session Footer");
        Assert.StartsWith("Footer settings updated for", result);
    }

    [Fact]
    public void SetSlideNumbering_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreatePresentation("test_session_numbering.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_slide_numbering", sessionId: sessionId, showSlideNumber: true, firstNumber: 10);
        Assert.StartsWith("Slide numbers", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(10, ppt.FirstSlideNumber);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("set_size", sessionId: "invalid_session", preset: "OnScreen16x9"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentation("test_path_setup.pptx");
        var pptPath2 = CreatePresentation("test_session_setup.pptx");
        var sessionId = OpenSession(pptPath2);
        _tool.Execute("set_slide_numbering", pptPath1, sessionId, showSlideNumber: true, firstNumber: 99);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(99, ppt.FirstSlideNumber);
    }

    #endregion
}
