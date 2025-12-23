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
        // Use the default first slide instead of AddEmptySlide to ensure shapes are properly saved
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task AddAnimation_ShouldAddAnimation()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_animation.pptx");

        // Find the correct shapeIndex for the added AutoShape (excluding placeholders)
        var correctShapeIndex = -1;
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var nonPlaceholderShapes = pptSlide.Shapes.Where(s => s.Placeholder == null).ToList();
            // If no non-placeholder shapes found, use all shapes
            if (nonPlaceholderShapes.Count == 0) nonPlaceholderShapes = pptSlide.Shapes.ToList();
            Assert.True(nonPlaceholderShapes.Count > 0,
                $"Should find at least one shape. Total shapes: {pptSlide.Shapes.Count}, Non-placeholder: {pptSlide.Shapes.Count(s => s.Placeholder == null)}");
            // The added shape should be the one with original coordinates (100, 100)
            for (var i = 0; i < nonPlaceholderShapes.Count; i++)
            {
                var s = nonPlaceholderShapes[i];
                if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                {
                    correctShapeIndex = pptSlide.Shapes.IndexOf(s);
                    break;
                }
            }

            if (correctShapeIndex < 0)
                correctShapeIndex =
                    pptSlide.Shapes.IndexOf(nonPlaceholderShapes[0]); // Fallback to first non-placeholder shape
        }

        Assert.True(correctShapeIndex >= 0, $"Should find at least one shape. Found shape index: {correctShapeIndex}");

        var outputPath = CreateTestFilePath("test_add_animation_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = correctShapeIndex,
            ["effectType"] = "Fade"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var sequence = slide.Timeline.MainSequence;
        Assert.True(sequence.Count > 0, "Slide should contain at least one animation effect");
    }

    [Fact]
    public async Task EditAnimation_ShouldModifyAnimation()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_edit_animation.pptx");

        // Find the correct shapeIndex for the added AutoShape (excluding placeholders)
        var correctShapeIndex = -1;
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            var nonPlaceholderShapes = slide.Shapes.Where(s => s.Placeholder == null).ToList();
            // If no non-placeholder shapes found, use all shapes
            if (nonPlaceholderShapes.Count == 0) nonPlaceholderShapes = slide.Shapes.ToList();
            Assert.True(nonPlaceholderShapes.Count > 0,
                $"Should find at least one shape. Total shapes: {slide.Shapes.Count}, Non-placeholder: {slide.Shapes.Count(s => s.Placeholder == null)}");
            // The added shape should be the one with original coordinates (100, 100)
            for (var i = 0; i < nonPlaceholderShapes.Count; i++)
            {
                var s = nonPlaceholderShapes[i];
                if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                {
                    correctShapeIndex = slide.Shapes.IndexOf(s);
                    break;
                }
            }

            if (correctShapeIndex < 0)
                correctShapeIndex =
                    slide.Shapes.IndexOf(nonPlaceholderShapes[0]); // Fallback to first non-placeholder shape

            var shape = slide.Shapes[correctShapeIndex];
            var sequence = slide.Timeline.MainSequence;
            sequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        Assert.True(correctShapeIndex >= 0, $"Should find at least one shape. Found shape index: {correctShapeIndex}");

        var outputPath = CreateTestFilePath("test_edit_animation_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = correctShapeIndex,
            ["effectType"] = "Fly",
            ["duration"] = 2.0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public async Task DeleteAnimation_ShouldDeleteAnimation()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_delete_animation.pptx");

        // Find the correct shapeIndex for the added AutoShape (excluding placeholders)
        var correctShapeIndex = -1;
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var nonPlaceholderShapes = pptSlide.Shapes.Where(s => s.Placeholder == null).ToList();
            // If no non-placeholder shapes found, use all shapes
            if (nonPlaceholderShapes.Count == 0) nonPlaceholderShapes = pptSlide.Shapes.ToList();
            Assert.True(nonPlaceholderShapes.Count > 0,
                $"Should find at least one shape. Total shapes: {pptSlide.Shapes.Count}, Non-placeholder: {pptSlide.Shapes.Count(s => s.Placeholder == null)}");
            // The added shape should be the one with original coordinates (100, 100)
            for (var i = 0; i < nonPlaceholderShapes.Count; i++)
            {
                var s = nonPlaceholderShapes[i];
                if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                {
                    correctShapeIndex = pptSlide.Shapes.IndexOf(s);
                    break;
                }
            }

            if (correctShapeIndex < 0)
                correctShapeIndex =
                    pptSlide.Shapes.IndexOf(nonPlaceholderShapes[0]); // Fallback to first non-placeholder shape

            var shape = pptSlide.Shapes[correctShapeIndex];
            var sequence = pptSlide.Timeline.MainSequence;
            sequence.AddEffect(shape, EffectType.Fade, EffectSubtype.None, EffectTriggerType.OnClick);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        Assert.True(correctShapeIndex >= 0, $"Should find at least one shape. Found shape index: {correctShapeIndex}");

        int animationsBefore;
        using (var tempPresentation = new Presentation(pptPath))
        {
            animationsBefore = tempPresentation.Slides[0].Timeline.MainSequence.Count;
        }

        Assert.True(animationsBefore > 0, "Animation should exist before deletion");

        var outputPath = CreateTestFilePath("test_delete_animation_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = correctShapeIndex
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var animationsAfter = presentation.Slides[0].Timeline.MainSequence.Count;
        Assert.True(animationsAfter < animationsBefore,
            $"Animation should be deleted. Before: {animationsBefore}, After: {animationsAfter}");
    }
}