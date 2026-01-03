using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Tests for the unified PptShapeTool (12 operations)
/// </summary>
public class PptShapeToolTests : TestBase
{
    private readonly PptShapeTool _tool;

    public PptShapeToolTests()
    {
        _tool = new PptShapeTool(SessionManager);
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

    #region General Tests

    [Fact]
    public void Delete_ShouldRemoveShape()
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

        _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: shapeIndex, outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        var shapesAfter = presentation.Slides[0].Shapes.Count;

        if (!IsEvaluationMode())
            Assert.True(shapesAfter < shapesBefore);
    }

    [Fact]
    public void Get_ShouldReturnAllShapes()
    {
        var pptPath = CreateTestPresentation("test_get_shapes.pptx");

        var result = _tool.Execute("get", pptPath, slideIndex: 0);

        Assert.NotNull(result);
        Assert.Contains("userShapeIndex", result);
        Assert.Contains("userShapeCount", result);
        Assert.Contains("totalCount", result);
    }

    [Fact]
    public void GetDetails_ShouldReturnShapeDetails()
    {
        var pptPath = CreateTestPresentation("test_get_details.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        Assert.True(shapeIndex >= 0);

        var result = _tool.Execute("get_details", pptPath, slideIndex: 0, shapeIndex: shapeIndex);

        Assert.Contains("position", result);
        Assert.Contains("size", result);
        Assert.Contains("fill", result);
        Assert.Contains("line", result);
        Assert.Contains("properties", result);
    }

    [Fact]
    public void Edit_ShouldModifyPosition()
    {
        var pptPath = CreateTestPresentation("test_edit_pos.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_pos_output.pptx");

        var result = _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: shapeIndex, x: 200, y: 200,
            outputPath: outputPath);

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
    public void Edit_WithRotation_ShouldRotateShape()
    {
        var pptPath = CreateTestPresentation("test_edit_rotation.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_rotation_output.pptx");

        _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rotation: 45, outputPath: outputPath);

        using var presentation = new Presentation(outputPath);
        var shape = presentation.Slides[0].Shapes[shapeIndex];
        if (!IsEvaluationMode())
            Assert.Equal(45, shape.Rotation, 1);
    }

    [SkippableFact]
    public void Edit_WithText_ShouldUpdateText()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestPresentation("test_edit_text.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_text_output.pptx");

        var result = _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: shapeIndex, text: "Updated Text",
            outputPath: outputPath);

        Assert.Contains("Text updated", result);
        using var ppt = new Presentation(outputPath);
        var shape = ppt.Slides[0].Shapes[shapeIndex] as IAutoShape;
        Assert.NotNull(shape?.TextFrame);
        Assert.Contains("Updated Text", shape.TextFrame.Text);
    }

    [Fact]
    public void SetFormat_ShouldSetFillColor()
    {
        var pptPath = CreateTestPresentation("test_set_format.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_set_format_output.pptx");

        var result = _tool.Execute("set_format", pptPath, slideIndex: 0, shapeIndex: shapeIndex, fillColor: "#FF0000",
            lineColor: "#0000FF", outputPath: outputPath);

        Assert.Contains("format updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void ClearFormat_ShouldClearFill()
    {
        var pptPath = CreateTestPresentation("test_clear_format.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_clear_format_output.pptx");

        var result = _tool.Execute("clear_format", pptPath, slideIndex: 0, shapeIndex: shapeIndex, clearFill: true,
            outputPath: outputPath);

        Assert.Contains("format cleared", result);
    }

    [Fact]
    public void Group_ShouldGroupShapes()
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

        var result = _tool.Execute("group", pptPath, slideIndex: 0, shapeIndices: [shapeIndex0, shapeIndex1],
            outputPath: outputPath);

        Assert.Contains("Grouped", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides[0].Shapes.Count > 0);
    }

    [Fact]
    public void Ungroup_ShouldUngroupShape()
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

        var result = _tool.Execute("ungroup", pptPath, slideIndex: 0, shapeIndex: groupShapeIndex,
            outputPath: outputPath);

        Assert.Contains("Ungrouped", result);
    }

    [Fact]
    public void Copy_ShouldCopyShapeToAnotherSlide()
    {
        var pptPath = CreateTestPresentation("test_copy.pptx");

        using (var pres = new Presentation(pptPath))
        {
            pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);
            pres.Save(pptPath, SaveFormat.Pptx);
        }

        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_copy_output.pptx");

        var result = _tool.Execute("copy", pptPath, shapeIndex: shapeIndex, fromSlide: 0, toSlide: 1,
            outputPath: outputPath);

        Assert.Contains("copied", result);
        using var presentation = new Presentation(outputPath);
        Assert.True(presentation.Slides[1].Shapes.Count > 0);
    }

    [Fact]
    public void Reorder_ShouldChangeZOrder()
    {
        var pptPath = CreateTestPresentationWithTwoShapes("test_reorder.pptx");

        int shapeIndex;
        using (var pres = new Presentation(pptPath))
        {
            var nonPlaceholder = pres.Slides[0].Shapes.Where(s => s.Placeholder == null).ToList();
            shapeIndex = pres.Slides[0].Shapes.IndexOf(nonPlaceholder[0]);
        }

        var outputPath = CreateTestFilePath("test_reorder_output.pptx");

        var result = _tool.Execute("reorder", pptPath, slideIndex: 0, shapeIndex: shapeIndex, toIndex: 0,
            outputPath: outputPath);

        Assert.Contains("Z-order", result);
    }

    [Fact]
    public void Align_ShouldAlignShapes()
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

        var result = _tool.Execute("align", pptPath, slideIndex: 0, shapeIndices: [shapeIndex0, shapeIndex1],
            align: "left", outputPath: outputPath);

        Assert.Contains("Aligned", result);
    }

    [Fact]
    public void Flip_ShouldFlipShape()
    {
        var pptPath = CreateTestPresentation("test_flip.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_flip_output.pptx");

        var result = _tool.Execute("flip", pptPath, slideIndex: 0, shapeIndex: shapeIndex, flipHorizontal: true,
            outputPath: outputPath);

        Assert.Contains("flipped", result);
    }

    [Fact]
    public void Flip_Vertical_ShouldFlipShape()
    {
        var pptPath = CreateTestPresentation("test_flip_v.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_flip_v_output.pptx");

        var result = _tool.Execute("flip", pptPath, slideIndex: 0, shapeIndex: shapeIndex, flipVertical: true,
            outputPath: outputPath);

        Assert.Contains("flipped", result);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_WithUnknownOperation_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");

        var ex = Assert.ThrowsAny<Exception>(() => _tool.Execute("unknown", pptPath, slideIndex: 0));
        if (ex is ArgumentException)
            Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void ExecuteAsync_WithInvalidSlideIndex_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");

        Assert.Throws<ArgumentException>(() => _tool.Execute("get", pptPath, slideIndex: 99));
    }

    [Fact]
    public void Edit_WithInvalidShapeIndex_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_invalid_shape.pptx");

        Assert.Throws<ArgumentException>(() => _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: 99, x: 100));
    }

    [Fact]
    public void ClearFormat_WithNoFlags_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_clear_no_flags.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("clear_format", pptPath, slideIndex: 0, shapeIndex: shapeIndex));
    }

    [Fact]
    public void Flip_WithNoDirection_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_flip_no_dir.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);

        Assert.Throws<ArgumentException>(() => _tool.Execute("flip", pptPath, slideIndex: 0, shapeIndex: shapeIndex));
    }

    [Fact]
    public void Group_WithLessThan2Shapes_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_group_one.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("group", pptPath, slideIndex: 0, shapeIndices: [shapeIndex]));
    }

    [Fact]
    public void Ungroup_WithNonGroupShape_ShouldThrowException()
    {
        var pptPath = CreateTestPresentation("test_ungroup_non_group.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);

        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("ungroup", pptPath, slideIndex: 0, shapeIndex: shapeIndex));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void Get_WithSessionId_ShouldReturnShapesFromMemory()
    {
        var pptPath = CreateTestPresentation("test_session_get.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        Assert.NotNull(result);
        Assert.Contains("userShapeIndex", result);
        Assert.Contains("totalCount", result);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_edit.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = 0;
        foreach (var s in ppt.Slides[0].Shapes)
            if (s.Placeholder == null)
            {
                shapeIndex = ppt.Slides[0].Shapes.IndexOf(s);
                break;
            }

        var result = _tool.Execute("edit", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex, x: 300, y: 300);
        Assert.Contains("edited", result);
        var shape = ppt.Slides[0].Shapes[shapeIndex];
        if (!IsEvaluationMode())
        {
            Assert.Equal(300, shape.X, 1);
            Assert.Equal(300, shape.Y, 1);
        }
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.Count;
        var shapeIndex = 0;
        foreach (var s in ppt.Slides[0].Shapes)
            if (s.Placeholder == null)
            {
                shapeIndex = ppt.Slides[0].Shapes.IndexOf(s);
                break;
            }

        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        if (!IsEvaluationMode()) Assert.True(ppt.Slides[0].Shapes.Count < initialCount);
    }

    [Fact]
    public void SetFormat_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_format.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = 0;
        foreach (var s in ppt.Slides[0].Shapes)
            if (s.Placeholder == null)
            {
                shapeIndex = ppt.Slides[0].Shapes.IndexOf(s);
                break;
            }

        var result = _tool.Execute("set_format", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            fillColor: "#00FF00");
        Assert.Contains("format updated", result);
        Assert.Contains("session", result);
    }

    #endregion
}