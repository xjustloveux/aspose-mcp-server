using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptAnimationToolTests : TestBase
{
    private readonly PptAnimationTool _tool = new();

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

    #region Error Handling Tests

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_operation.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion

    #region Add Animation Tests

    [Fact]
    public async Task AddAnimation_ShouldAddAnimation()
    {
        var pptPath = CreateTestPresentation("test_add_animation.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_animation_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["effectType"] = "Fade"
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides[0].Timeline.MainSequence.Count > 0);
    }

    [Fact]
    public async Task AddAnimation_WithSubtype_ShouldApplySubtype()
    {
        var pptPath = CreateTestPresentation("test_add_animation_subtype.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_animation_subtype_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["effectType"] = "Fly",
            ["effectSubtype"] = "FromBottom"
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        var sequence = presentation.Slides[0].Timeline.MainSequence;
        Assert.True(sequence.Count > 0);
        Assert.Equal(EffectType.Fly, sequence[0].Type);
    }

    [Fact]
    public async Task AddAnimation_WithTriggerType_ShouldApplyTrigger()
    {
        var pptPath = CreateTestPresentation("test_add_animation_trigger.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_animation_trigger_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["effectType"] = "Fade",
            ["triggerType"] = "AfterPrevious"
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        var sequence = presentation.Slides[0].Timeline.MainSequence;
        Assert.True(sequence.Count > 0);
        Assert.Equal(EffectTriggerType.AfterPrevious, sequence[0].Timing.TriggerType);
    }

    [Fact]
    public async Task AddAnimation_InvalidShapeIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_add_animation_invalid.pptx");
        var outputPath = CreateTestFilePath("test_add_animation_invalid_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 999,
            ["effectType"] = "Fade"
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion

    #region Edit Animation Tests

    [Fact]
    public async Task EditAnimation_ShouldModifyAnimation()
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["effectType"] = "Fly",
            ["duration"] = 2.0
        };

        await _tool.ExecuteAsync(arguments);

        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task EditAnimation_WithAnimationIndex_ShouldModifySpecificAnimation()
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["animationIndex"] = 0,
            ["duration"] = 3.0
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(2, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public async Task EditAnimation_InvalidAnimationIndex_ShouldThrow()
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["animationIndex"] = 999,
            ["effectType"] = "Fly"
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion

    #region Delete Animation Tests

    [Fact]
    public async Task DeleteAnimation_ShouldDeleteAnimation()
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
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(0, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public async Task DeleteAnimation_WithAnimationIndex_ShouldDeleteSpecificAnimation()
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
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["animationIndex"] = 0
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(1, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    [Fact]
    public async Task DeleteAnimation_AllFromSlide_ShouldClearSequence()
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
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        Assert.Equal(0, presentation.Slides[0].Timeline.MainSequence.Count);
    }

    #endregion
}