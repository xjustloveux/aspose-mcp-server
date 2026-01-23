using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Animation;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptAnimationTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptAnimationToolTests : PptTestBase
{
    private readonly PptAnimationTool _tool;

    public PptAnimationToolTests()
    {
        _tool = new PptAnimationTool(SessionManager);
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

    private static int FindShapeIndex(string pptPath)
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

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddAnimation()
    {
        var pptPath = CreatePresentationWithShape("test_add.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_output.pptx");
        var result = _tool.Execute("add", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fade",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Animation", data.Message);
        Assert.Contains("added", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides[0].Timeline.MainSequence.Count > 0);
    }

    [Fact]
    public void Edit_ShouldModifyAnimation()
    {
        var pptPath = CreatePresentationWithAnimation("test_edit.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_output.pptx");
        var result = _tool.Execute("edit", 0, pptPath, shapeIndex: shapeIndex, effectType: "Fly", duration: 2.0f,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Animation updated", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Delete_ShouldDeleteAnimationForShape()
    {
        var pptPath = CreatePresentationWithAnimation("test_delete.pptx");
        var outputPath = CreateTestFilePath("test_delete_output.pptx");
        var result = _tool.Execute("delete", 0, pptPath, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message, StringComparison.OrdinalIgnoreCase);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(0, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public void Get_WithAnimations_ShouldReturnAnimationList()
    {
        var pptPath = CreatePresentationWithAnimation("test_get.pptx");
        var result = _tool.Execute("get", 0, pptPath);
        var data = GetResultData<GetAnimationsResult>(result);
        Assert.True(data.TotalAnimationsOnSlide > 0);
        Assert.Contains(data.Animations, a => a.EffectType == "Fade");
    }

    [Fact]
    public void Get_EmptySlide_ShouldReturnEmptyList()
    {
        var pptPath = CreatePresentationWithShape("test_get_empty.pptx");
        var result = _tool.Execute("get", 0, pptPath);
        var data = GetResultData<GetAnimationsResult>(result);
        Assert.Equal(0, data.TotalAnimationsOnSlide);
        Assert.Empty(data.Animations);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentationWithShape($"test_case_add_{operation}.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, 0, pptPath, shapeIndex: shapeIndex, effectType: "Fade",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("added", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithShape("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", 0, pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentationWithShape("test_session_add.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Timeline.MainSequence.Count;
        var result = _tool.Execute("add", 0, sessionId: sessionId, shapeIndex: shapeIndex, effectType: "Fade");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Animation", data.Message);
        Assert.True(ppt.Slides[0].Timeline.MainSequence.Count > initialCount);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreatePresentationWithAnimation("test_session_edit.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("edit", 0, sessionId: sessionId, shapeIndex: shapeIndex, animationIndex: 0,
            duration: 3.0f);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Animation updated", data.Message);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        Assert.Equal(3.0f, ppt.Slides[0].Timeline.MainSequence[0].Timing.Duration, 1);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithAnimation("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Timeline.MainSequence.Count;
        var result = _tool.Execute("delete", 0, sessionId: sessionId);
        var data = GetResultData<SuccessResult>(result);
        Assert.Contains("deleted", data.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(ppt.Slides[0].Timeline.MainSequence.Count < initialCount);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Get_WithSessionId_ShouldReturnAnimations()
    {
        var pptPath = CreatePresentationWithAnimation("test_session_get.pptx", EffectType.Appear);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", 0, sessionId: sessionId);
        var data = GetResultData<GetAnimationsResult>(result);
        Assert.Equal(1, data.TotalAnimationsOnSlide);
        Assert.Contains(data.Animations, a => a.EffectType == "Appear");
        var output = GetResultOutput<GetAnimationsResult>(result);
        Assert.True(output.IsSession);
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
        var data = GetResultData<GetAnimationsResult>(result);
        Assert.Contains(data.Animations, a => a.EffectType == "Zoom");
    }

    #endregion
}
