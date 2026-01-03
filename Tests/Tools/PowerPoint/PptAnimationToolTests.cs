using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptAnimationToolTests : TestBase
{
    private readonly PptAnimationTool _tool;

    public PptAnimationToolTests()
    {
        _tool = new PptAnimationTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private int FindShapeIndex(string pptPath)
    {
        using var ppt = new Presentation(pptPath);
        var slide = ppt.Slides[0];
        var nonPlaceholderShapes = slide.Shapes.Where(s => s.Placeholder == null).ToList();
        if (nonPlaceholderShapes.Count == 0) nonPlaceholderShapes = slide.Shapes.ToList();
        foreach (var s in nonPlaceholderShapes)
            if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                return slide.Shapes.IndexOf(s);
        return nonPlaceholderShapes.Count > 0 ? slide.Shapes.IndexOf(nonPlaceholderShapes[0]) : 0;
    }

    #region General Tests

    #region Add Animation Tests

    [Fact]
    public void AddAnimation_ShouldAddAnimation()
    {
        var pptPath = CreateTestPresentation("test_add_animation.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_animation_output.pptx");

        _tool.Execute("add", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fade", outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides[0].Timeline.MainSequence.Count > 0);
    }

    [Fact]
    public void AddAnimation_WithSubtype_ShouldApplySubtype()
    {
        var pptPath = CreateTestPresentation("test_add_animation_subtype.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_animation_subtype_output.pptx");

        _tool.Execute("add", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fly", effectSubtype: "FromBottom",
            outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        var sequence = presentation.Slides[0].Timeline.MainSequence;
        Assert.True(sequence.Count > 0);
        Assert.Equal(EffectType.Fly, sequence[0].Type);
    }

    [Fact]
    public void AddAnimation_WithTriggerType_ShouldApplyTrigger()
    {
        var pptPath = CreateTestPresentation("test_add_animation_trigger.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_animation_trigger_output.pptx");

        _tool.Execute("add", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fade", triggerType: "AfterPrevious",
            outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        var sequence = presentation.Slides[0].Timeline.MainSequence;
        Assert.True(sequence.Count > 0);
        Assert.Equal(EffectTriggerType.AfterPrevious, sequence[0].Timing.TriggerType);
    }

    [Fact]
    public void AddAnimation_InvalidShapeIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_add_animation_invalid.pptx");
        var outputPath = CreateTestFilePath("test_add_animation_invalid_output.pptx");

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", 0, pptPath, shapeIndex: 999, effectType: "Fade", outputPath: outputPath));
    }

    #endregion

    #region Edit Animation Tests

    [Fact]
    public void EditAnimation_ShouldModifyAnimation()
    {
        var pptPath = CreateTestPresentation("test_edit_animation.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            ppt.Slides[0].Timeline.MainSequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None,
                EffectTriggerType.OnClick);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_edit_animation_output.pptx");

        _tool.Execute("edit", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fly", duration: 2.0f,
            outputPath: outputPath);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void EditAnimation_WithAnimationIndex_ShouldModifySpecificAnimation()
    {
        var pptPath = CreateTestPresentation("test_edit_animation_index.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            var seq = ppt.Slides[0].Timeline.MainSequence;
            seq.AddEffect(shape, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
            seq.AddEffect(shape, EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_edit_animation_index_output.pptx");

        _tool.Execute("edit", 0, pptPath, shapeIndex: shapeIndex, animationIndex: 0, duration: 3.0f,
            outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(2, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public void EditAnimation_InvalidAnimationIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_edit_animation_invalid_index.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            ppt.Slides[0].Timeline.MainSequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None,
                EffectTriggerType.OnClick);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_edit_animation_invalid_index_output.pptx");

        Assert.Throws<ArgumentException>(() => _tool.Execute("edit", 0, pptPath, shapeIndex: shapeIndex,
            animationIndex: 999, effectType: "Fly", outputPath: outputPath));
    }

    #endregion

    #region Get Animation Tests

    [Fact]
    public void GetAnimations_EmptySlide_ShouldReturnEmptyList()
    {
        var pptPath = CreateTestPresentation("test_get_no_animation.pptx");

        var result = _tool.Execute("get", 0, pptPath);

        Assert.Contains("\"totalAnimationsOnSlide\": 0", result);
        Assert.Contains("\"animations\": []", result);
    }

    [Fact]
    public void GetAnimations_WithAnimations_ShouldReturnAnimationList()
    {
        var pptPath = CreateTestPresentation("test_get_animations.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            var sequence = ppt.Slides[0].Timeline.MainSequence;
            sequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
            sequence.AddEffect(shape, EffectType.Fly, EffectSubtype.Bottom, EffectTriggerType.AfterPrevious);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", 0, pptPath);

        Assert.Contains("\"totalAnimationsOnSlide\": 2", result);
        Assert.Contains("\"effectType\": \"Fade\"", result);
        Assert.Contains("\"effectType\": \"Fly\"", result);
        Assert.Contains("\"effectSubtype\": \"Bottom\"", result);
        Assert.Contains("\"triggerType\": \"OnClick\"", result);
        Assert.Contains("\"triggerType\": \"AfterPrevious\"", result);
    }

    [Fact]
    public void GetAnimations_FilterByShapeIndex_ShouldReturnFilteredList()
    {
        var pptPath = CreateTestPresentation("test_get_animation_filter.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        // Add another shape and animations
        using (var ppt = new Presentation(pptPath))
        {
            var slide = ppt.Slides[0];
            var shape1 = slide.Shapes[shapeIndex];
            var shape2 = slide.Shapes.AddAutoShape(ShapeType.Ellipse, 300, 100, 100, 100);
            var sequence = slide.Timeline.MainSequence;
            sequence.AddEffect(shape1, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
            sequence.AddEffect(shape2, EffectType.Zoom, EffectSubtype.None, EffectTriggerType.OnClick);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", 0, pptPath, shapeIndex: shapeIndex);

        Assert.Contains("\"totalAnimationsOnSlide\": 2", result);
        Assert.Contains("\"effectType\": \"Fade\"", result);
        Assert.DoesNotContain("\"effectType\": \"Zoom\"", result);
    }

    [Fact]
    public void GetAnimations_WithDurationAndDelay_ShouldReturnTimingInfo()
    {
        var pptPath = CreateTestPresentation("test_get_animation_timing.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            var sequence = ppt.Slides[0].Timeline.MainSequence;
            var effect = sequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
            effect.Timing.Duration = 2.5f;
            effect.Timing.TriggerDelayTime = 1.0f;
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", 0, pptPath);

        Assert.Contains("\"duration\":", result);
        Assert.Contains("\"delay\":", result);
    }

    [Fact]
    public void GetAnimations_WithSessionId_ShouldReturnAnimations()
    {
        var pptPath = CreateTestPresentation("test_get_animation_session.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            ppt.Slides[0].Timeline.MainSequence.AddEffect(shape, EffectType.Appear, EffectSubtype.None,
                EffectTriggerType.OnClick);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);

        var result = _tool.Execute("get", 0, sessionId: sessionId);

        Assert.Contains("\"totalAnimationsOnSlide\": 1", result);
        Assert.Contains("\"effectType\": \"Appear\"", result);
    }

    #endregion

    #region Delete Animation Tests

    [Fact]
    public void DeleteAnimation_ShouldDeleteAnimation()
    {
        var pptPath = CreateTestPresentation("test_delete_animation.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            ppt.Slides[0].Timeline.MainSequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None,
                EffectTriggerType.OnClick);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_animation_output.pptx");

        _tool.Execute("delete", 0, pptPath, shapeIndex: shapeIndex, outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(0, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public void DeleteAnimation_WithAnimationIndex_ShouldDeleteSpecificAnimation()
    {
        var pptPath = CreateTestPresentation("test_delete_animation_index.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            var sequence = ppt.Slides[0].Timeline.MainSequence;
            sequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
            sequence.AddEffect(shape, EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_animation_index_output.pptx");

        _tool.Execute("delete", 0, pptPath, shapeIndex: shapeIndex, animationIndex: 0, outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(1, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public void DeleteAnimation_AllFromSlide_ShouldClearSequence()
    {
        var pptPath = CreateTestPresentation("test_delete_all_animation.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            var sequence = ppt.Slides[0].Timeline.MainSequence;
            sequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
            sequence.AddEffect(shape, EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_all_animation_output.pptx");

        _tool.Execute("delete", 0, pptPath, outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(0, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    #endregion

    #endregion

    #region Exception Tests

    [Fact]
    public void Execute_UnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_operation.pptx");

        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", 0, pptPath));
    }

    [Fact]
    public void AddAnimation_WithDefaultEffectType_ShouldUseFade()
    {
        var pptPath = CreateTestPresentation("test_add_default_effect.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_default_effect_output.pptx");

        // Act - effectType defaults to "Fade" when not provided
        var result = _tool.Execute("add", 0, pptPath, shapeIndex: shapeIndex, outputPath: outputPath);
        Assert.Contains("Animation", result);
        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        using var presentation = new Presentation(outputPath);
        var sequence = presentation.Slides[0].Timeline.MainSequence;
        Assert.True(sequence.Count > 0);
        Assert.Equal(EffectType.Fade, sequence[0].Type);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddAnimation_WithSessionId_ShouldVerifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_verify_animation.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("add", 0, sessionId: sessionId, shapeIndex: shapeIndex, effectType: "Fade");
        Assert.NotNull(result);
        Assert.Contains("animation", result, StringComparison.OrdinalIgnoreCase);

        // Verify animation was added in memory
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        var animationCount = slide.Timeline.MainSequence.Count;
        Assert.True(animationCount > 0, "Animation should be added to the slide");
    }

    [Fact]
    public void AddAnimation_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add_animation.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var sessionId = OpenSession(pptPath);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Timeline.MainSequence.Count;
        var result = _tool.Execute("add", 0, sessionId: sessionId, shapeIndex: shapeIndex, effectType: "Fade");
        Assert.Contains("Animation", result);
        Assert.Contains("added", result);
        Assert.True(ppt.Slides[0].Timeline.MainSequence.Count > initialCount);
    }

    [Fact]
    public void DeleteAnimation_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_delete_animation.pptx");
        var shapeIndex = FindShapeIndex(pptPath);

        // Add animation before opening session
        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            ppt.Slides[0].Timeline.MainSequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None,
                EffectTriggerType.OnClick);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var pptSession = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = pptSession.Slides[0].Timeline.MainSequence.Count;
        var result = _tool.Execute("delete", 0, sessionId: sessionId);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(pptSession.Slides[0].Timeline.MainSequence.Count < initialCount);
    }

    #endregion
}