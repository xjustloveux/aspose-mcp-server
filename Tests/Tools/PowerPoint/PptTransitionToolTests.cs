using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Results;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Transition;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptTransitionTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptTransitionToolTests : PptTestBase
{
    private readonly PptTransitionTool _tool;

    public PptTransitionToolTests()
    {
        _tool = new PptTransitionTool(SessionManager);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void SetTransition_ShouldSetTransition()
    {
        var pptPath = CreatePresentation("test_set_transition.pptx");
        var outputPath = CreateTestFilePath("test_set_transition_output.pptx");
        _tool.Execute("set", pptPath, slideIndex: 0, transitionType: "Fade", advanceAfterSeconds: 1.5,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.Equal(TransitionType.Fade, slide.SlideShowTransition.Type);
        Assert.Equal(1500u, slide.SlideShowTransition.AdvanceAfterTime);
    }

    [Fact]
    public void GetTransition_ShouldReturnTransitionInfo()
    {
        var pptPath = CreatePresentation("test_get_transition.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            slide.SlideShowTransition.AdvanceAfterTime = 2000;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var data = GetResultData<GetTransitionResult>(result);
        Assert.Equal("Fade", data.Type);
        Assert.True(data.HasTransition);
    }

    [Fact]
    public void GetTransition_WithNoTransition_ShouldReturnHasTransitionFalse()
    {
        var pptPath = CreatePresentation("test_get_no_transition.pptx");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var data = GetResultData<GetTransitionResult>(result);
        Assert.False(data.HasTransition);
        Assert.Equal("None", data.Type);
    }

    [Fact]
    public void DeleteTransition_ShouldRemoveTransition()
    {
        var pptPath = CreatePresentation("test_delete_transition.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_transition_output.pptx");
        _tool.Execute("delete", pptPath, slideIndex: 0, outputPath: outputPath);
        using var resultPresentation = new Presentation(outputPath);
        Assert.Equal(TransitionType.None, resultPresentation.Slides[0].SlideShowTransition.Type);
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
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, transitionType: "Fade", outputPath: outputPath);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void SetTransition_WithSessionId_ShouldSetInMemory()
    {
        var pptPath = CreatePresentation("test_session_set.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, transitionType: "Push",
            advanceAfterSeconds: 2.5);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.NotNull(SessionManager.GetSessionStatus(sessionId));

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        Assert.Equal(TransitionType.Push, slide.SlideShowTransition.Type);
    }

    [Fact]
    public void GetTransition_WithSessionId_ShouldReturnTransition()
    {
        var pptPath = CreateTestFilePath("test_session_get.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        var data = GetResultData<GetTransitionResult>(result);
        Assert.Equal("Fade", data.Type);
        Assert.True(data.HasTransition);
        Assert.NotNull(SessionManager.GetSessionStatus(sessionId));
    }

    [Fact]
    public void DeleteTransition_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestFilePath("test_session_delete.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Wipe;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0);
        Assert.IsType<FinalizedResult<SuccessResult>>(result);
        Assert.NotNull(SessionManager.GetSessionStatus(sessionId));

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(TransitionType.None, ppt.Slides[0].SlideShowTransition.Type);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("get", sessionId: "invalid_session_id", slideIndex: 0));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestFilePath("test_path_trans.pptx");
        using (var pres1 = new Presentation())
        {
            pres1.Slides[0].SlideShowTransition.Type = TransitionType.Fade;
            pres1.Save(pptPath1, SaveFormat.Pptx);
        }

        var pptPath2 = CreateTestFilePath("test_session_trans.pptx");
        using (var pres2 = new Presentation())
        {
            pres2.Slides[0].SlideShowTransition.Type = TransitionType.Push;
            pres2.Save(pptPath2, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId, slideIndex: 0);
        var data = GetResultData<GetTransitionResult>(result);
        Assert.Equal("Push", data.Type);
    }

    #endregion
}
