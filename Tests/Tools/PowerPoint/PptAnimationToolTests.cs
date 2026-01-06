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

    private string CreatePresentationWithAnimation(string fileName, EffectType effectType = EffectType.Fade)
    {
        var filePath = CreateTestFilePath(fileName);
        using var ppt = new Presentation();
        var slide = ppt.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        slide.Timeline.MainSequence.AddEffect(shape, effectType, EffectSubtype.None, EffectTriggerType.OnClick);
        ppt.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithMultipleAnimations(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var ppt = new Presentation();
        var slide = ppt.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        var seq = slide.Timeline.MainSequence;
        seq.AddEffect(shape, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
        seq.AddEffect(shape, EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious);
        ppt.Save(filePath, SaveFormat.Pptx);
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

    #region General

    [Fact]
    public void Add_ShouldAddAnimation()
    {
        var pptPath = CreateTestPresentation("test_add.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_output.pptx");
        var result = _tool.Execute("add", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fade",
            outputPath: outputPath);
        Assert.StartsWith("Animation", result);
        Assert.Contains("added", result); // Verify action was completed
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides[0].Timeline.MainSequence.Count > 0);
    }

    [Fact]
    public void Add_WithDefaultEffectType_ShouldUseFade()
    {
        var pptPath = CreateTestPresentation("test_add_default.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_default_output.pptx");
        _tool.Execute("add", 0, pptPath, shapeIndex: shapeIndex, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var sequence = presentation.Slides[0].Timeline.MainSequence;
        Assert.Equal(EffectType.Fade, sequence[0].Type);
    }

    [Fact]
    public void Add_WithSubtype_ShouldApplySubtype()
    {
        var pptPath = CreateTestPresentation("test_add_subtype.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_subtype_output.pptx");
        _tool.Execute("add", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fly", effectSubtype: "FromBottom",
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var sequence = presentation.Slides[0].Timeline.MainSequence;
        Assert.Equal(EffectType.Fly, sequence[0].Type);
    }

    [Fact]
    public void Add_WithTriggerType_ShouldApplyTrigger()
    {
        var pptPath = CreateTestPresentation("test_add_trigger.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_trigger_output.pptx");
        _tool.Execute("add", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fade", triggerType: "AfterPrevious",
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var sequence = presentation.Slides[0].Timeline.MainSequence;
        Assert.Equal(EffectTriggerType.AfterPrevious, sequence[0].Timing.TriggerType);
    }

    [Fact]
    public void Edit_ShouldModifyAnimation()
    {
        var pptPath = CreatePresentationWithAnimation("test_edit.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_output.pptx");
        var result = _tool.Execute("edit", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fly", duration: 2.0f,
            outputPath: outputPath);
        Assert.StartsWith("Animation updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Edit_WithAnimationIndex_ShouldModifySpecificAnimation()
    {
        var pptPath = CreatePresentationWithMultipleAnimations("test_edit_index.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_index_output.pptx");
        _tool.Execute("edit", 0, pptPath, shapeIndex: shapeIndex, animationIndex: 0, duration: 3.0f,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(2, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public void Edit_WithDurationAndDelay_ShouldApplyTiming()
    {
        var pptPath = CreatePresentationWithAnimation("test_edit_timing.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_timing_output.pptx");
        _tool.Execute("edit", 0, pptPath, shapeIndex: shapeIndex, animationIndex: 0, duration: 2.5f, delay: 1.0f,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var effect = presentation.Slides[0].Timeline.MainSequence[0];
        Assert.Equal(2.5f, effect.Timing.Duration, 1);
        Assert.Equal(1.0f, effect.Timing.TriggerDelayTime, 1);
    }

    [Fact]
    public void Delete_ShouldDeleteAnimationForShape()
    {
        var pptPath = CreatePresentationWithAnimation("test_delete.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_delete_output.pptx");
        var result = _tool.Execute("delete", 0, pptPath, shapeIndex: shapeIndex, outputPath: outputPath);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(0, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public void Delete_WithAnimationIndex_ShouldDeleteSpecificAnimation()
    {
        var pptPath = CreatePresentationWithMultipleAnimations("test_delete_index.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_delete_index_output.pptx");
        _tool.Execute("delete", 0, pptPath, shapeIndex: shapeIndex, animationIndex: 0, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(1, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public void Delete_AllFromSlide_ShouldClearSequence()
    {
        var pptPath = CreatePresentationWithMultipleAnimations("test_delete_all.pptx");
        var outputPath = CreateTestFilePath("test_delete_all_output.pptx");
        _tool.Execute("delete", 0, pptPath, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(0, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public void Get_EmptySlide_ShouldReturnEmptyList()
    {
        var pptPath = CreateTestPresentation("test_get_empty.pptx");
        var result = _tool.Execute("get", 0, pptPath);
        Assert.Contains("\"totalAnimationsOnSlide\": 0", result);
        Assert.Contains("\"animations\": []", result);
    }

    [Fact]
    public void Get_WithAnimations_ShouldReturnAnimationList()
    {
        var pptPath = CreatePresentationWithMultipleAnimations("test_get.pptx");
        var result = _tool.Execute("get", 0, pptPath);
        Assert.Contains("\"totalAnimationsOnSlide\": 2", result);
        Assert.Contains("\"effectType\": \"Fade\"", result);
        Assert.Contains("\"effectType\": \"Appear\"", result);
    }

    [Fact]
    public void Get_FilterByShapeIndex_ShouldReturnFilteredList()
    {
        var pptPath = CreateTestPresentation("test_get_filter.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
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
    public void Get_ShouldIncludeTimingInfo()
    {
        var pptPath = CreateTestPresentation("test_get_timing.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        using (var ppt = new Presentation(pptPath))
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            var effect = ppt.Slides[0].Timeline.MainSequence
                .AddEffect(shape, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
            effect.Timing.Duration = 2.5f;
            effect.Timing.TriggerDelayTime = 1.0f;
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", 0, pptPath);
        Assert.Contains("\"duration\":", result);
        Assert.Contains("\"delay\":", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_add_{operation}.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, 0, pptPath, shapeIndex: shapeIndex, effectType: "Fade",
            outputPath: outputPath);
        Assert.Contains("added", result); // Verify action was completed
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var pptPath = CreatePresentationWithAnimation($"test_case_edit_{operation}.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.pptx");
        var result = _tool.Execute(operation, 0, pptPath, shapeIndex: shapeIndex, duration: 2.0f,
            outputPath: outputPath);
        Assert.StartsWith("Animation updated", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var pptPath = CreatePresentationWithAnimation($"test_case_delete_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_delete_{operation}_output.pptx");
        var result = _tool.Execute(operation, 0, pptPath, outputPath: outputPath);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_get_{operation}.pptx");
        var result = _tool.Execute(operation, 0, pptPath);
        Assert.Contains("totalAnimationsOnSlide", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", 0, pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_no_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", 0, pptPath, effectType: "Fade"));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void Add_WithInvalidShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_invalid_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", 0, pptPath, shapeIndex: 999, effectType: "Fade"));
        Assert.Contains("shape", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Edit_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithAnimation("test_edit_no_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("edit", 0, pptPath, effectType: "Fly"));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidAnimationIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithAnimation("test_edit_invalid_anim.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", 0, pptPath, shapeIndex: shapeIndex, animationIndex: 999, effectType: "Fly"));
        Assert.Contains("animationIndex", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidAnimationIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithAnimation("test_delete_invalid_anim.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("delete", 0, pptPath, shapeIndex: shapeIndex, animationIndex: 999));
        Assert.Contains("animationIndex", ex.Message);
    }

    [Fact]
    public void Execute_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", 99, pptPath));
        Assert.Contains("slide", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Session

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Timeline.MainSequence.Count;
        var result = _tool.Execute("add", 0, sessionId: sessionId, shapeIndex: shapeIndex, effectType: "Fade");
        Assert.StartsWith("Animation", result);
        Assert.Contains("session", result); // Verify session was used
        Assert.True(ppt.Slides[0].Timeline.MainSequence.Count > initialCount);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreatePresentationWithAnimation("test_session_edit.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("edit", 0, sessionId: sessionId, shapeIndex: shapeIndex, animationIndex: 0,
            duration: 3.0f);
        Assert.StartsWith("Animation updated", result);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(3.0f, ppt.Slides[0].Timeline.MainSequence[0].Timing.Duration, 1);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithAnimation("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Timeline.MainSequence.Count;
        var result = _tool.Execute("delete", 0, sessionId: sessionId);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        Assert.True(ppt.Slides[0].Timeline.MainSequence.Count < initialCount);
    }

    [Fact]
    public void Get_WithSessionId_ShouldReturnAnimations()
    {
        var pptPath = CreatePresentationWithAnimation("test_session_get.pptx", EffectType.Appear);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", 0, sessionId: sessionId);
        Assert.Contains("\"totalAnimationsOnSlide\": 1", result);
        Assert.Contains("\"effectType\": \"Appear\"", result);
    }

    [Fact]
    public void Get_WithSessionId_AfterAdd_ShouldReflectChanges()
    {
        var pptPath = CreateTestPresentation("test_session_get_after_add.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var sessionId = OpenSession(pptPath);
        _tool.Execute("add", 0, sessionId: sessionId, shapeIndex: shapeIndex, effectType: "Zoom");
        var result = _tool.Execute("get", 0, sessionId: sessionId);
        Assert.Contains("\"totalAnimationsOnSlide\": 1", result);
        Assert.Contains("\"effectType\": \"Zoom\"", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", 0, sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithAnimation("test_path_anim.pptx");
        var pptPath2 = CreatePresentationWithAnimation("test_session_anim.pptx", EffectType.Zoom);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", 0, pptPath1, sessionId);
        Assert.Contains("\"effectType\": \"Zoom\"", result);
        Assert.DoesNotContain("\"effectType\": \"Fade\"", result);
    }

    #endregion
}