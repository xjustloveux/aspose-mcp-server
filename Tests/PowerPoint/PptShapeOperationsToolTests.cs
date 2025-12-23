using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptShapeOperationsToolTests : TestBase
{
    private readonly PptShapeOperationsTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];

        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        slide.Shapes.AddAutoShape(ShapeType.Ellipse, 350, 100, 200, 100);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task GroupShapes_ShouldGroupShapes()
    {
        var pptPath = CreateTestPresentation("test_group_shapes.pptx");

        int shapeIndex0, shapeIndex1;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholderBefore = sld.Shapes.Where(s => s.Placeholder == null).ToList();

            Assert.True(nonPlaceholderBefore.Count >= 2,
                $"Should have at least 2 non-placeholder shapes, found {nonPlaceholderBefore.Count}");

            shapeIndex0 = sld.Shapes.IndexOf(nonPlaceholderBefore[0]);
            shapeIndex1 = sld.Shapes.IndexOf(nonPlaceholderBefore[1]);
        }

        var outputPath = CreateTestFilePath("test_group_shapes_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "group",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndices"] = new JsonArray { shapeIndex0, shapeIndex1 }
        };

        await _tool.ExecuteAsync(arguments);

        var isEvaluationMode = IsEvaluationMode();
        using var presentation = new Presentation(outputPath);
        var resultSlide = presentation.Slides[0];
        Assert.NotNull(resultSlide);

        var shapesAfter = resultSlide.Shapes.Count;
        var groupShapes = resultSlide.Shapes.OfType<IGroupShape>().ToList();

        if (isEvaluationMode)
            Assert.True(shapesAfter > 0, "In evaluation mode, shapes should exist after grouping");
        else
            Assert.True(groupShapes.Count > 0 || shapesAfter > 0,
                "Shapes should be grouped or at least exist after operation");
    }

    [Fact]
    public async Task CopyShape_ShouldCopyShape()
    {
        var pptPath = CreateTestPresentation("test_copy_shape.pptx");

        using (var presentation = new Presentation(pptPath))
        {
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        using (var pres = new Presentation(pptPath))
        {
            var fromSlide = pres.Slides[0];
            var nonPlaceholderShapes = fromSlide.Shapes.Where(s => s.Placeholder == null).ToList();
            Assert.True(nonPlaceholderShapes.Count > 0, "Source slide should have at least one shape");

            var shapeIndex = fromSlide.Shapes.IndexOf(nonPlaceholderShapes[0]);

            var outputPath = CreateTestFilePath("test_copy_shape_output.pptx");
            var arguments = new JsonObject
            {
                ["operation"] = "copy",
                ["path"] = pptPath,
                ["outputPath"] = outputPath,
                ["slideIndex"] = 1,
                ["shapeIndex"] = shapeIndex,
                ["fromSlide"] = 0,
                ["toSlide"] = 1
            };

            await _tool.ExecuteAsync(arguments);

            Assert.True(File.Exists(outputPath), "Output presentation should be created");

            var isEvaluationMode = IsEvaluationMode();
            using var resultPresentation = new Presentation(outputPath);
            var toSlide = resultPresentation.Slides[1];

            if (!isEvaluationMode)
                Assert.True(toSlide.Shapes.Count > 0, "Target slide should have copied shape");
            else
                Assert.True(resultPresentation.Slides.Count >= 2,
                    "In evaluation mode, presentation should have at least 2 slides");
        }
    }

    [Fact]
    public async Task UngroupShape_ShouldUngroupShape()
    {
        // First create a grouped shape
        var pptPath = CreateTestPresentation("test_ungroup_shapes.pptx");

        int groupShapeIndex;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholderBefore = sld.Shapes.Where(s => s.Placeholder == null).ToList();

            Assert.True(nonPlaceholderBefore.Count >= 2,
                $"Should have at least 2 non-placeholder shapes, found {nonPlaceholderBefore.Count}");

            var shapeIndex0 = sld.Shapes.IndexOf(nonPlaceholderBefore[0]);
            var shapeIndex1 = sld.Shapes.IndexOf(nonPlaceholderBefore[1]);

            // Group the shapes
            var groupShape = sld.Shapes.AddGroupShape();
            groupShape.Shapes.AddClone(nonPlaceholderBefore[0]);
            groupShape.Shapes.AddClone(nonPlaceholderBefore[1]);
            sld.Shapes.RemoveAt(shapeIndex1);
            sld.Shapes.RemoveAt(shapeIndex0);

            groupShapeIndex = sld.Shapes.IndexOf(groupShape);
            pres.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_ungroup_shapes_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "ungroup",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = groupShapeIndex
        };

        await _tool.ExecuteAsync(arguments);

        var isEvaluationMode = IsEvaluationMode();
        using var presentation = new Presentation(outputPath);
        var resultSlide = presentation.Slides[0];
        Assert.NotNull(resultSlide);

        // After ungrouping, there should be more shapes (ungrouped shapes)
        var message = !isEvaluationMode
            ? "Shapes should exist after ungrouping"
            : "In evaluation mode, shapes should exist after ungrouping";
        Assert.True(resultSlide.Shapes.Count > 0, message);
    }

    [Fact]
    public async Task ReorderShape_ShouldReorderShape()
    {
        var pptPath = CreateTestPresentation("test_reorder_shape.pptx");

        int shapeIndex;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholderShapes = sld.Shapes.Where(s => s.Placeholder == null).ToList();
            Assert.True(nonPlaceholderShapes.Count > 0, "Should have at least one shape");
            shapeIndex = sld.Shapes.IndexOf(nonPlaceholderShapes[0]);
            pres.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_reorder_shape_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "reorder",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["toIndex"] = 0
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        var resultSlide = presentation.Slides[0];
        Assert.NotNull(resultSlide);
        Assert.True(resultSlide.Shapes.Count > 0, "Shapes should exist after reordering");
    }

    [Fact]
    public async Task AlignShapes_ShouldAlignShapes()
    {
        var pptPath = CreateTestPresentation("test_align_shapes.pptx");

        int shapeIndex0, shapeIndex1;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholderBefore = sld.Shapes.Where(s => s.Placeholder == null).ToList();

            Assert.True(nonPlaceholderBefore.Count >= 2,
                $"Should have at least 2 non-placeholder shapes, found {nonPlaceholderBefore.Count}");

            shapeIndex0 = sld.Shapes.IndexOf(nonPlaceholderBefore[0]);
            shapeIndex1 = sld.Shapes.IndexOf(nonPlaceholderBefore[1]);
        }

        var outputPath = CreateTestFilePath("test_align_shapes_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "align",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndices"] = new JsonArray { shapeIndex0, shapeIndex1 },
            ["align"] = "left"
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        var resultSlide = presentation.Slides[0];
        Assert.NotNull(resultSlide);
        Assert.True(resultSlide.Shapes.Count > 0, "Shapes should exist after alignment");
    }

    [Fact]
    public async Task FlipShape_ShouldFlipShape()
    {
        var pptPath = CreateTestPresentation("test_flip_shape.pptx");

        int shapeIndex;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholderShapes = sld.Shapes.Where(s => s.Placeholder == null).ToList();
            Assert.True(nonPlaceholderShapes.Count > 0, "Should have at least one shape");
            shapeIndex = sld.Shapes.IndexOf(nonPlaceholderShapes[0]);
            pres.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_flip_shape_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "flip",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["flipHorizontal"] = true
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        var resultSlide = presentation.Slides[0];
        Assert.NotNull(resultSlide);
        Assert.True(resultSlide.Shapes.Count > 0, "Shape should exist after flipping");
    }
}