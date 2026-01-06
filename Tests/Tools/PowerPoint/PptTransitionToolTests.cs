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

    private string CreateTestPresentation(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General

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
    public void SetTransition_WithoutAdvanceAfterSeconds_ShouldNotSetAutoAdvance()
    {
        var pptPath = CreateTestPresentation("test_set_no_advance.pptx");
        var outputPath = CreateTestFilePath("test_set_no_advance_output.pptx");
        _tool.Execute("set", pptPath, slideIndex: 0, transitionType: "Wipe", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.Equal(TransitionType.Wipe, slide.SlideShowTransition.Type);
    }

    [Theory]
    [InlineData("Fade", TransitionType.Fade)]
    [InlineData("Push", TransitionType.Push)]
    [InlineData("Wipe", TransitionType.Wipe)]
    [InlineData("Split", TransitionType.Split)]
    [InlineData("Circle", TransitionType.Circle)]
    [InlineData("Diamond", TransitionType.Diamond)]
    public void SetTransition_WithVariousTypes_ShouldSetCorrectType(string typeStr, TransitionType expected)
    {
        var pptPath = CreateTestPresentation($"test_set_{typeStr}.pptx");
        var outputPath = CreateTestFilePath($"test_set_{typeStr}_output.pptx");
        _tool.Execute("set", pptPath, slideIndex: 0, transitionType: typeStr, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(expected, presentation.Slides[0].SlideShowTransition.Type);
    }

    [Fact]
    public void SetTransition_OnSecondSlide_ShouldSetOnCorrectSlide()
    {
        var pptPath = CreateTestPresentation("test_set_slide2.pptx", 3);
        var outputPath = CreateTestFilePath("test_set_slide2_output.pptx");
        _tool.Execute("set", pptPath, slideIndex: 1, transitionType: "Push", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(TransitionType.None, presentation.Slides[0].SlideShowTransition.Type);
        Assert.Equal(TransitionType.Push, presentation.Slides[1].SlideShowTransition.Type);
    }

    [Fact]
    public void GetTransition_ShouldReturnTransitionInfo()
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
        Assert.Contains("\"type\": \"Fade\"", result);
        Assert.Contains("\"hasTransition\": true", result);
        Assert.Contains("\"advanceAfterSeconds\": 2", result);
    }

    [Fact]
    public void GetTransition_WithNoTransition_ShouldReturnHasTransitionFalse()
    {
        var pptPath = CreateTestPresentation("test_get_no_transition.pptx");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        Assert.Contains("\"hasTransition\": false", result);
        Assert.Contains("\"type\": \"None\"", result);
    }

    [Fact]
    public void GetTransition_ShouldReturnJsonWithCorrectFields()
    {
        var pptPath = CreateTestPresentation("test_get_fields.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            presentation.Slides[0].SlideShowTransition.Type = TransitionType.Wipe;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        Assert.Contains("\"slideIndex\"", result);
        Assert.Contains("\"hasTransition\"", result);
        Assert.Contains("\"type\"", result);
        Assert.Contains("\"speed\"", result);
        Assert.Contains("\"advanceOnClick\"", result);
    }

    [Fact]
    public void DeleteTransition_ShouldRemoveTransition()
    {
        var pptPath = CreateTestPresentation("test_delete_transition.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            slide.SlideShowTransition.AdvanceAfterTime = 1500;
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
    public void DeleteTransition_OnSlideWithoutTransition_ShouldSucceed()
    {
        var pptPath = CreateTestPresentation("test_delete_no_trans.pptx");
        var outputPath = CreateTestFilePath("test_delete_no_trans_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Transition removed", result);
    }

    [Theory]
    [InlineData("SET")]
    [InlineData("Set")]
    [InlineData("set")]
    public void Operation_ShouldBeCaseInsensitive_Set(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_set_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_set_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, transitionType: "Fade", outputPath: outputPath);
        Assert.StartsWith("Transition 'Fade' set", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_get_{operation}.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0);
        Assert.Contains("\"slideIndex\"", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_del_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_del_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Transition removed", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void SetTransition_WithInvalidType_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_type.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 0, transitionType: "InvalidType"));
        Assert.Contains("Invalid transition type", ex.Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void SetTransition_WithEmptyOrNullType_ShouldThrowArgumentException(string? transitionType)
    {
        var pptPath = CreateTestPresentation("test_empty_type.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 0, transitionType: transitionType));
        Assert.Contains("transitionType is required", ex.Message);
    }

    [Fact]
    public void SetTransition_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: 999, transitionType: "Fade"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void SetTransition_WithNegativeSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_neg_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("set", pptPath, slideIndex: -1, transitionType: "Fade"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void GetTransition_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_get_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("get", pptPath, slideIndex: 999));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void DeleteTransition_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_del_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", pptPath, slideIndex: 999));
        Assert.Contains("out of range", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void SetTransition_WithSessionId_ShouldSetInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_set.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("set", sessionId: sessionId, slideIndex: 0, transitionType: "Push",
            advanceAfterSeconds: 2.5);
        Assert.StartsWith("Transition 'Push' set for slide 0", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        Assert.Equal(TransitionType.Push, slide.SlideShowTransition.Type);
        Assert.Equal(2500u, slide.SlideShowTransition.AdvanceAfterTime);
    }

    [Fact]
    public void GetTransition_WithSessionId_ShouldReturnTransition()
    {
        var pptPath = CreateTestFilePath("test_session_get.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            slide.SlideShowTransition.Type = TransitionType.Fade;
            slide.SlideShowTransition.AdvanceAfterTime = 3000;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        Assert.Contains("\"type\": \"Fade\"", result);
        Assert.Contains("\"hasTransition\": true", result);
    }

    [Fact]
    public void DeleteTransition_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestFilePath("test_session_delete.pptx");
        using (var presentation = new Presentation())
        {
            var slideToSetup = presentation.Slides[0];
            slideToSetup.SlideShowTransition.Type = TransitionType.Wipe;
            slideToSetup.SlideShowTransition.AdvanceAfterTime = 1500;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var pptBefore = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(TransitionType.Wipe, pptBefore.Slides[0].SlideShowTransition.Type);

        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0);
        Assert.StartsWith("Transition removed", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        Assert.Equal(TransitionType.None, slide.SlideShowTransition.Type);
        Assert.Equal(0u, slide.SlideShowTransition.AdvanceAfterTime);
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
        Assert.Contains("\"type\": \"Push\"", result);
        Assert.DoesNotContain("\"type\": \"Fade\"", result);
    }

    #endregion
}