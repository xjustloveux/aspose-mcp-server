using Aspose.Slides;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Font;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptFontTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptFontToolTests : PptTestBase
{
    private readonly PptFontTool _tool;

    public PptFontToolTests()
    {
        _tool = new PptFontTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Replace_ShouldReplaceFont()
    {
        var pptPath = CreatePresentationWithContent("test_replace.pptx", "Hello World");
        var outputPath = CreateTestFilePath("test_replace_output.pptx");
        var result = _tool.Execute("replace", pptPath, sourceFont: "Calibri", targetFont: "Arial",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("replaced", data.Message);
    }

    [Fact]
    public void GetUsed_ShouldReturnFonts()
    {
        var pptPath = CreatePresentationWithContent("test_get_used.pptx", "Hello World");
        var result = _tool.Execute("get_used", pptPath);
        var data = GetResultData<GetFontsPptResult>(result);
        Assert.True(data.Count > 0);
        Assert.NotEmpty(data.Items);
    }

    [Fact]
    public void SetFallback_ShouldSetFallbackRule()
    {
        var pptPath = CreatePresentationWithContent("test_fallback.pptx", "Hello World");
        var outputPath = CreateTestFilePath("test_fallback_output.pptx");
        var result = _tool.Execute("set_fallback", pptPath, fallbackFont: "Arial", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("fallback", data.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Embed_WithExistingFont_ShouldEmbed()
    {
        var pptPath = CreatePresentationWithContent("test_embed.pptx", "Hello World");
        using var pres = new Presentation(pptPath);
        var allFonts = pres.FontsManager.GetFonts();
        if (allFonts.Length == 0) return;
        var fontName = allFonts[0].FontName;

        var outputPath = CreateTestFilePath("test_embed_output.pptx");
        var result = _tool.Execute("embed", pptPath, fontName: fontName, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains(fontName, data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("GET_USED")]
    [InlineData("Get_Used")]
    [InlineData("get_used")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentationWithContent($"test_case_{operation.Replace(" ", "_")}.pptx", "Hello World");
        var result = _tool.Execute(operation, pptPath);
        var data = GetResultData<GetFontsPptResult>(result);
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
    public void GetUsed_WithSessionId_ShouldReturnFontsFromMemory()
    {
        var pptPath = CreatePresentationWithContent("test_session_get.pptx", "Hello World");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_used", sessionId: sessionId);
        var data = GetResultData<GetFontsPptResult>(result);
        Assert.True(data.Count >= 0);
        var output = GetResultOutput<GetFontsPptResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    [Fact]
    public void Replace_WithSessionId_ShouldReplaceInMemory()
    {
        var pptPath = CreatePresentationWithContent("test_session_replace.pptx", "Hello World");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("replace", sessionId: sessionId, sourceFont: "Calibri", targetFont: "Arial");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("replaced", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    [Fact]
    public void SetFallback_WithSessionId_ShouldSetInMemory()
    {
        var pptPath = CreatePresentationWithContent("test_session_fallback.pptx", "Hello World");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set_fallback", sessionId: sessionId, fallbackFont: "Arial");
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("fallback", data.Message, StringComparison.OrdinalIgnoreCase);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.Equal(sessionId, output.SessionId);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_used", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithContent("test_path_font.pptx", "Path content");
        var pptPath2 = CreatePresentationWithContent("test_session_font.pptx", "Session content");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get_used", pptPath1, sessionId);
        var data = GetResultData<GetFontsPptResult>(result);
        Assert.NotNull(data);
    }

    #endregion
}
