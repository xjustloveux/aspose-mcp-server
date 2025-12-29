using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

/// <summary>
///     Tests for the unified PptShapeTool (12 operations)
/// </summary>
public class PptShapeToolTests : TestBase
{
    private readonly PptShapeTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreateTestPresentationWithTwoShapes(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        slide.Shapes.AddAutoShape(ShapeType.Ellipse, 350, 100, 200, 100);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private static int FindNonPlaceholderShapeIndex(string pptPath)
    {
        using var ppt = new Presentation(pptPath);
        var slide = ppt.Slides[0];
        var nonPlaceholderShapes = slide.Shapes.Where(s => s.Placeholder == null).ToList();
        if (nonPlaceholderShapes.Count == 0) return -1;

        foreach (var s in nonPlaceholderShapes)
            if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                return slide.Shapes.IndexOf(s);

        return slide.Shapes.IndexOf(nonPlaceholderShapes[0]);
    }

    #region Basic Operations - Delete

    [Fact]
    public async Task Delete_ShouldRemoveShape()
    {
        var pptPath = CreateTestPresentation("test_delete.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        Assert.True(shapeIndex >= 0);

        int shapesBefore;
        using (var pres = new Presentation(pptPath))
        {
            shapesBefore = pres.Slides[0].Shapes.Count;
        }

        var outputPath = CreateTestFilePath("test_delete_output.pptx");
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
        var shapesAfter = presentation.Slides[0].Shapes.Count;

        if (!IsEvaluationMode())
            Assert.True(shapesAfter < shapesBefore);
    }

    #endregion

    #region Basic Operations - Get

    [Fact]
    public async Task Get_ShouldReturnAllShapes()
    {
        var pptPath = CreateTestPresentation("test_get_shapes.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.NotNull(result);
        Assert.Contains("userShapeIndex", result);
        Assert.Contains("userShapeCount", result);
        Assert.Contains("totalCount", result);
    }

    [Fact]
    public async Task GetDetails_ShouldReturnShapeDetails()
    {
        var pptPath = CreateTestPresentation("test_get_details.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        Assert.True(shapeIndex >= 0);

        var arguments = new JsonObject
        {
            ["operation"] = "get_details",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("position", result);
        Assert.Contains("size", result);
        Assert.Contains("fill", result);
        Assert.Contains("line", result);
        Assert.Contains("properties", result);
    }

    #endregion

    #region Edit Operations

    [Fact]
    public async Task Edit_ShouldModifyPosition()
    {
        var pptPath = CreateTestPresentation("test_edit_pos.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_pos_output.pptx");

        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["x"] = 200,
            ["y"] = 200
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("edited", result);
        using var presentation = new Presentation(outputPath);
        var shape = presentation.Slides[0].Shapes[shapeIndex];
        if (!IsEvaluationMode())
        {
            Assert.Equal(200, shape.X, 1);
            Assert.Equal(200, shape.Y, 1);
        }
    }

    [Fact]
    public async Task Edit_WithRotation_ShouldRotateShape()
    {
        var pptPath = CreateTestPresentation("test_edit_rotation.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_rotation_output.pptx");

        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["rotation"] = 45
        };

        await _tool.ExecuteAsync(arguments);

        using var presentation = new Presentation(outputPath);
        var shape = presentation.Slides[0].Shapes[shapeIndex];
        if (!IsEvaluationMode())
            Assert.Equal(45, shape.Rotation, 1);
    }

    [Fact]
    public async Task Edit_WithText_ShouldUpdateText()
    {
        var pptPath = CreateTestPresentation("test_edit_text.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_text_output.pptx");

        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["text"] = "Updated Text"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Text updated", result);
    }

    [Fact]
    public async Task SetFormat_ShouldSetFillColor()
    {
        var pptPath = CreateTestPresentation("test_set_format.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_set_format_output.pptx");

        var arguments = new JsonObject
        {
            ["operation"] = "set_format",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["fillColor"] = "#FF0000",
            ["lineColor"] = "#0000FF"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("format updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task ClearFormat_ShouldClearFill()
    {
        var pptPath = CreateTestPresentation("test_clear_format.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_clear_format_output.pptx");

        var arguments = new JsonObject
        {
            ["operation"] = "clear_format",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["clearFill"] = true
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("format cleared", result);
    }

    #endregion

    #region Advanced Operations - Group/Ungroup

    [Fact]
    public async Task Group_ShouldGroupShapes()
    {
        var pptPath = CreateTestPresentationWithTwoShapes("test_group.pptx");

        int shapeIndex0, shapeIndex1;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholder = sld.Shapes.Where(s => s.Placeholder == null).ToList();
            Assert.True(nonPlaceholder.Count >= 2);
            shapeIndex0 = sld.Shapes.IndexOf(nonPlaceholder[0]);
            shapeIndex1 = sld.Shapes.IndexOf(nonPlaceholder[1]);
        }

        var outputPath = CreateTestFilePath("test_group_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "group",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndices"] = new JsonArray { shapeIndex0, shapeIndex1 }
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Grouped", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides[0].Shapes.Count > 0);
    }

    [Fact]
    public async Task Ungroup_ShouldUngroupShape()
    {
        var pptPath = CreateTestPresentationWithTwoShapes("test_ungroup.pptx");

        int groupShapeIndex;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholder = sld.Shapes.Where(s => s.Placeholder == null).ToList();
            var idx0 = sld.Shapes.IndexOf(nonPlaceholder[0]);
            var idx1 = sld.Shapes.IndexOf(nonPlaceholder[1]);

            var groupShape = sld.Shapes.AddGroupShape();
            groupShape.Shapes.AddClone(nonPlaceholder[0]);
            groupShape.Shapes.AddClone(nonPlaceholder[1]);
            sld.Shapes.RemoveAt(idx1);
            sld.Shapes.RemoveAt(idx0);

            groupShapeIndex = sld.Shapes.IndexOf(groupShape);
            pres.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_ungroup_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "ungroup",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = groupShapeIndex
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Ungrouped", result);
    }

    #endregion

    #region Advanced Operations - Copy/Reorder

    [Fact]
    public async Task Copy_ShouldCopyShapeToAnotherSlide()
    {
        var pptPath = CreateTestPresentation("test_copy.pptx");

        using (var pres = new Presentation(pptPath))
        {
            pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);
            pres.Save(pptPath, SaveFormat.Pptx);
        }

        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_copy_output.pptx");

        var arguments = new JsonObject
        {
            ["operation"] = "copy",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["fromSlide"] = 0,
            ["toSlide"] = 1,
            ["shapeIndex"] = shapeIndex
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("copied", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides[1].Shapes.Count > 0);
    }

    [Fact]
    public async Task Reorder_ShouldChangeZOrder()
    {
        var pptPath = CreateTestPresentationWithTwoShapes("test_reorder.pptx");

        int shapeIndex;
        using (var pres = new Presentation(pptPath))
        {
            var nonPlaceholder = pres.Slides[0].Shapes.Where(s => s.Placeholder == null).ToList();
            shapeIndex = pres.Slides[0].Shapes.IndexOf(nonPlaceholder[0]);
        }

        var outputPath = CreateTestFilePath("test_reorder_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "reorder",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["toIndex"] = 0
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Z-order", result);
    }

    #endregion

    #region Advanced Operations - Align/Flip

    [Fact]
    public async Task Align_ShouldAlignShapes()
    {
        var pptPath = CreateTestPresentationWithTwoShapes("test_align.pptx");

        int shapeIndex0, shapeIndex1;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholder = sld.Shapes.Where(s => s.Placeholder == null).ToList();
            shapeIndex0 = sld.Shapes.IndexOf(nonPlaceholder[0]);
            shapeIndex1 = sld.Shapes.IndexOf(nonPlaceholder[1]);
        }

        var outputPath = CreateTestFilePath("test_align_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "align",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndices"] = new JsonArray { shapeIndex0, shapeIndex1 },
            ["align"] = "left"
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("Aligned", result);
    }

    [Fact]
    public async Task Flip_ShouldFlipShape()
    {
        var pptPath = CreateTestPresentation("test_flip.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_flip_output.pptx");

        var arguments = new JsonObject
        {
            ["operation"] = "flip",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["flipHorizontal"] = true
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("flipped", result);
    }

    [Fact]
    public async Task Flip_Vertical_ShouldFlipShape()
    {
        var pptPath = CreateTestPresentation("test_flip_v.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_flip_v_output.pptx");

        var arguments = new JsonObject
        {
            ["operation"] = "flip",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex,
            ["flipVertical"] = true
        };

        var result = await _tool.ExecuteAsync(arguments);

        Assert.Contains("flipped", result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public async Task ExecuteAsync_WithUnknownOperation_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ExecuteAsync_WithInvalidSlideIndex_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 99
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task Edit_WithInvalidShapeIndex_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_invalid_shape.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 99,
            ["x"] = 100
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ClearFormat_WithNoFlags_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_clear_no_flags.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var arguments = new JsonObject
        {
            ["operation"] = "clear_format",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task Flip_WithNoDirection_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_flip_no_dir.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var arguments = new JsonObject
        {
            ["operation"] = "flip",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task Group_WithLessThan2Shapes_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_group_one.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var arguments = new JsonObject
        {
            ["operation"] = "group",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndices"] = new JsonArray { shapeIndex }
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task Ungroup_WithNonGroupShape_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_ungroup_non_group.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var arguments = new JsonObject
        {
            ["operation"] = "ungroup",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = shapeIndex
        };

        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    #endregion
}