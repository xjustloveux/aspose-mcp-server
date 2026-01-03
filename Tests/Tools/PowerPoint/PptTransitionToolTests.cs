using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptTransitionToolTests : TestBase
{
    private readonly PptTransitionTool _tool;

    public PptTransitionToolTests()
    {
        _tool = new PptTransitionTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void SetTransition_ShouldSetTransition()
    {
        var pptPath = CreateTestPresentation("test_set_transition.pptx");
        var outputPath = CreateTestFilePath("test_set_transition_output.pptx");
        _tool.Execute("set", pptPath, slideIndex: 0, transitionType: "Fade", advanceAfterSeconds: 1.5,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.Equal(TransitionType.Fade, slide.SlideShowTransition.Type);
        Assert.Equal(1500u, slide.SlideShowTransition.AdvanceAfterTime);
    }

    [Fact]
    public void GetTransition_ShouldReturnTransition()
    {
        var pptPath = CreateTestPresentation("test_get_transition.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            slide.SlideShowTransition.AdvanceAfterTime = 2000;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        Assert.NotNull(result);
        Assert.Contains("\"type\": \"Fade\"", result);
        Assert.Contains("\"hasTransition\": true", result);
        Assert.Contains("\"advanceAfterSeconds\": 2", result);
    }

    [Fact]
    public void DeleteTransition_ShouldDeleteTransition()
    {
        var pptPath = CreateTestPresentation("test_delete_transition.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_transition_output.pptx");
        _tool.Execute("delete", pptPath, slideIndex: 0, outputPath: outputPath);
        using var resultPresentation = new Presentation(outputPath);
        var resultSlide = resultPresentation.Slides[0];
        Assert.Equal(TransitionType.None, resultSlide.SlideShowTransition.Type);
        Assert.Equal(0u, resultSlide.SlideShowTransition.AdvanceAfterTime);
    }

    [Fact]
    public void SetTransition_WithEnumTryParse_ShouldSupportAllTypes()
    {
        var pptPath = CreateTestPresentation("test_set_transition_push.pptx");
        var outputPath = CreateTestFilePath("test_set_transition_push_output.pptx");
        _tool.Execute("set", pptPath, slideIndex: 0, transitionType: "Push", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.Equal(TransitionType.Push, slide.SlideShowTransition.Type);
    }

    [Fact]
    public void SetTransition_WithoutAdvanceAfterSeconds_ShouldNotSetAutoAdvance()
    {
        var pptPath = CreateTestPresentation("test_set_transition_no_advance.pptx");
        var outputPath = CreateTestFilePath("test_set_transition_no_advance_output.pptx");
        _tool.Execute("set", pptPath, slideIndex: 0, transitionType: "Wipe", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.Equal(TransitionType.Wipe, slide.SlideShowTransition.Type);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath, slideIndex: 0));
    }

    [Fact]
    public void SetTransition_WithInvalidType_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_set_transition_invalid.pptx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 0, transitionType: "InvalidType"));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetTransition_WithSessionId_ShouldReturnTransition()
    {
        var pptPath = CreateTestFilePath("test_session_get_transition.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            slide.SlideShowTransition.AdvanceAfterTime = 3000;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        Assert.NotNull(result);
        Assert.Contains("\"type\": \"Fade\"", result);
        Assert.Contains("\"hasTransition\": true", result);
    }

    [Fact]
    public void SetTransition_WithSessionId_ShouldSetInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set_transition.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, transitionType: "Push",
            advanceAfterSeconds: 2.5);
        Assert.Contains("Transition 'Push' set for slide 0", result);

        // Verify in-memory changes
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        Assert.Equal(TransitionType.Push, slide.SlideShowTransition.Type);
        Assert.Equal(2500u, slide.SlideShowTransition.AdvanceAfterTime);
    }

    [Fact]
    public void DeleteTransition_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestFilePath("test_session_delete_transition.pptx");
        using (var presentation = new Presentation())
        {
            var slideToSetup = presentation.Slides[0];
            slideToSetup.SlideShowTransition.Type = TransitionType.Wipe;
            slideToSetup.SlideShowTransition.AdvanceAfterTime = 1500;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);

        // Verify transition exists before delete
        var pptBefore = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(TransitionType.Wipe, pptBefore.Slides[0].SlideShowTransition.Type);
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0);
        Assert.Contains("Transition removed", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        Assert.Equal(TransitionType.None, slide.SlideShowTransition.Type);
        Assert.Equal(0u, slide.SlideShowTransition.AdvanceAfterTime);
    }

    #endregion
}