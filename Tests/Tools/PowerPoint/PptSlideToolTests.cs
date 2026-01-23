using Aspose.Slides;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Slide;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptSlideTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptSlideToolTests : PptTestBase
{
    private readonly PptSlideTool _tool;

    public PptSlideToolTests()
    {
        _tool = new PptSlideTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddNewSlideAndPersistToFile()
    {
        var pptPath = CreatePresentation("test_add.pptx", 2);
        var outputPath = CreateTestFilePath("test_add_output.pptx");
        var result = _tool.Execute("add", pptPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide added", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(3, presentation.Slides.Count);
    }

    [Fact]
    public void Delete_ShouldDeleteSlideAndPersistToFile()
    {
        var pptPath = CreatePresentation("test_delete.pptx", 3);
        var outputPath = CreateTestFilePath("test_delete_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 1, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide 1 deleted", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(2, presentation.Slides.Count);
    }

    [Fact]
    public void GetInfo_ShouldReturnSlidesInfoFromFile()
    {
        var pptPath = CreatePresentation("test_get_info.pptx", 3);
        var result = _tool.Execute("get_info", pptPath);
        var data = GetResultData<GetSlidesInfoResult>(result);
        Assert.Equal(3, data.Count);
        Assert.NotNull(data.Slides);
    }

    [Fact]
    public void Move_ShouldMoveSlideAndPersistToFile()
    {
        var pptPath = CreatePresentation("test_move.pptx", 3);
        var outputPath = CreateTestFilePath("test_move_output.pptx");
        var result = _tool.Execute("move", pptPath, fromIndex: 2, toIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide moved", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(3, presentation.Slides.Count);
    }

    [Fact]
    public void Duplicate_ShouldDuplicateSlideAndPersistToFile()
    {
        var pptPath = CreatePresentation("test_duplicate.pptx", 2);
        var outputPath = CreateTestFilePath("test_duplicate_output.pptx");
        var result = _tool.Execute("duplicate", pptPath, slideIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide 0 duplicated", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(3, presentation.Slides.Count);
    }

    [Fact]
    public void Hide_ShouldHideSlideAndPersistToFile()
    {
        var pptPath = CreatePresentation("test_hide.pptx");
        var outputPath = CreateTestFilePath("test_hide_output.pptx");
        var result = _tool.Execute("hide", pptPath, slideIndex: 0, hidden: true, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Set", data.Message);
        Assert.Contains("hidden=True", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Clear_ShouldClearSlideContentAndPersistToFile()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Clear verification may not work correctly in evaluation mode");
        var pptPath = CreatePresentationWithShape("test_clear.pptx");
        var outputPath = CreateTestFilePath("test_clear_output.pptx");
        var result = _tool.Execute("clear", pptPath, slideIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Cleared all shapes", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Slides[0].Shapes);
    }

    [Fact]
    public void Edit_ShouldEditSlideAndPersistToFile()
    {
        var pptPath = CreatePresentation("test_edit.pptx");
        var outputPath = CreateTestFilePath("test_edit_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide 0 updated", data.Message);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide added", data.Message);
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
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides.Count;
        var result = _tool.Execute("add", sessionId: sessionId);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide added", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        Assert.Equal(initialCount + 1, ppt.Slides.Count);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentation("test_session_delete.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 1);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide 1 deleted", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        Assert.Equal(2, ppt.Slides.Count);
    }

    [Fact]
    public void GetInfo_WithSessionId_ShouldReturnInfo()
    {
        var pptPath = CreatePresentation("test_session_info.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_info", sessionId: sessionId);
        var data = GetResultData<GetSlidesInfoResult>(result);
        Assert.True(data.Count >= 0);
        var output = GetResultOutput<GetSlidesInfoResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Move_WithSessionId_ShouldMoveInMemory()
    {
        var pptPath = CreatePresentation("test_session_move.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("move", sessionId: sessionId, fromIndex: 2, toIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide moved", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Duplicate_WithSessionId_ShouldDuplicateInMemory()
    {
        var pptPath = CreatePresentation("test_session_dup.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides.Count;
        var result = _tool.Execute("duplicate", sessionId: sessionId, slideIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide 0 duplicated", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        Assert.Equal(initialCount + 1, ppt.Slides.Count);
    }

    [Fact]
    public void Hide_WithSessionId_ShouldHideInMemory()
    {
        var pptPath = CreatePresentation("test_session_hide.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("hide", sessionId: sessionId, slideIndex: 0, hidden: true);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Set", data.Message);
        Assert.Contains("hidden=True", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var pptPath = CreatePresentationWithShape("test_session_clear.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("clear", sessionId: sessionId, slideIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Cleared", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        Assert.Empty(ppt.Slides[0].Shapes);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var pptPath = CreatePresentation("test_session_edit.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("edit", sessionId: sessionId, slideIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Slide 0 updated", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_info", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithSessionIdButNoSessionManager_ShouldThrowInvalidOperationException()
    {
        var toolWithoutSession = new PptSlideTool();
        var ex = Assert.Throws<InvalidOperationException>(() =>
            toolWithoutSession.Execute("get_info", sessionId: "any_session"));
        Assert.Contains("Session management is not enabled", ex.Message);
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentation("test_path_ppt.pptx");
        var pptPath2 = CreatePresentation("test_session_ppt.pptx", 5);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get_info", pptPath1, sessionId);
        var data = GetResultData<GetSlidesInfoResult>(result);
        Assert.Equal(5, data.Count);
    }

    #endregion
}
