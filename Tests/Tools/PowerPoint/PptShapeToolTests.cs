using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

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

    private string CreatePresentationWithTwoShapes(string fileName)
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
        for (var i = 0; i < slide.Shapes.Count; i++)
            if (slide.Shapes[i].Placeholder == null)
                return i;
        return -1;
    }

    private static int FindNonPlaceholderShapeIndex(ISlide slide)
    {
        for (var i = 0; i < slide.Shapes.Count; i++)
            if (slide.Shapes[i].Placeholder == null)
                return i;
        return -1;
    }

    #region General

    [Fact]
    public void Get_ShouldReturnAllShapes()
    {
        var pptPath = CreateTestPresentation("test_get.pptx");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        Assert.Contains("\"userShapeIndex\":", result);
        Assert.Contains("\"userShapeCount\":", result);
        Assert.Contains("\"totalCount\":", result);
    }

    [Fact]
    public void GetDetails_ShouldReturnShapeDetails()
    {
        var pptPath = CreateTestPresentation("test_get_details.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var result = _tool.Execute("get_details", pptPath, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.Contains("\"position\":", result);
        Assert.Contains("\"size\":", result);
        Assert.Contains("\"fill\":", result);
        Assert.Contains("\"line\":", result);
        Assert.Contains("\"properties\":", result);
    }

    [Fact]
    public void Delete_ShouldRemoveShape()
    {
        var pptPath = CreateTestPresentation("test_delete.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        int shapesBefore;
        using (var pres = new Presentation(pptPath))
        {
            shapesBefore = pres.Slides[0].Shapes.Count;
        }

        var outputPath = CreateTestFilePath("test_delete_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: shapeIndex, outputPath: outputPath);
        Assert.StartsWith("Shape", result);
        Assert.Contains("deleted from slide", result);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        if (!IsEvaluationMode())
            Assert.True(presentation.Slides[0].Shapes.Count < shapesBefore);
        else
            Assert.NotNull(presentation.Slides[0]);
    }

    [Fact]
    public void Edit_ShouldModifyPosition()
    {
        var pptPath = CreateTestPresentation("test_edit_pos.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_pos_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: shapeIndex, x: 200, y: 200,
            outputPath: outputPath);
        Assert.StartsWith("Shape", result);
        Assert.Contains("edited", result);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        var shape = presentation.Slides[0].Shapes[shapeIndex];
        if (!IsEvaluationMode())
        {
            Assert.Equal(200, shape.X, 1);
            Assert.Equal(200, shape.Y, 1);
        }
        else
        {
            Assert.NotNull(shape);
        }
    }

    [Fact]
    public void Edit_WithRotation_ShouldRotateShape()
    {
        var pptPath = CreateTestPresentation("test_edit_rotation.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_rotation_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: shapeIndex, rotation: 45,
            outputPath: outputPath);
        Assert.StartsWith("Shape", result);
        Assert.Contains("edited", result);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        if (!IsEvaluationMode())
            Assert.Equal(45, presentation.Slides[0].Shapes[shapeIndex].Rotation, 1);
        else
            Assert.NotNull(presentation.Slides[0].Shapes[shapeIndex]);
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
        Assert.StartsWith("Shape", result);
        Assert.Contains("edited", result);
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
        Assert.StartsWith("Shape", result);
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
        Assert.StartsWith("Shape", result);
        Assert.Contains("format cleared", result);
    }

    [Fact]
    public void Group_ShouldGroupShapes()
    {
        var pptPath = CreatePresentationWithTwoShapes("test_group.pptx");
        int idx0, idx1;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholder = sld.Shapes.Where(s => s.Placeholder == null).ToList();
            idx0 = sld.Shapes.IndexOf(nonPlaceholder[0]);
            idx1 = sld.Shapes.IndexOf(nonPlaceholder[1]);
        }

        var outputPath = CreateTestFilePath("test_group_output.pptx");
        var result = _tool.Execute("group", pptPath, slideIndex: 0, shapeIndices: [idx0, idx1], outputPath: outputPath);
        Assert.StartsWith("Grouped", result);
    }

    [Fact]
    public void Ungroup_ShouldUngroupShape()
    {
        var pptPath = CreatePresentationWithTwoShapes("test_ungroup.pptx");
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
        Assert.StartsWith("Ungrouped", result);
    }

    [Fact]
    public void Copy_ShouldCopyShapeToAnotherSlide()
    {
        var pptPath = CreateTestPresentation("test_copy.pptx");
        float sourceWidth, sourceHeight;
        using (var pres = new Presentation(pptPath))
        {
            pres.Slides.AddEmptySlide(pres.LayoutSlides[0]);
            var sourceShape = pres.Slides[0].Shapes.FirstOrDefault(s => s.Placeholder == null)!;
            sourceWidth = sourceShape.Width;
            sourceHeight = sourceShape.Height;
            pres.Save(pptPath, SaveFormat.Pptx);
        }

        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_copy_output.pptx");
        var result = _tool.Execute("copy", pptPath, shapeIndex: shapeIndex, fromSlide: 0, toSlide: 1,
            outputPath: outputPath);
        Assert.StartsWith("Shape", result);
        Assert.Contains("copied from slide", result);
        using var presentation = new Presentation(outputPath);
        var copiedShapes = presentation.Slides[1].Shapes.Where(s => s.Placeholder == null).ToList();
        Assert.True(copiedShapes.Count > 0, "Shape should be copied to slide 1");
        var copiedShape = copiedShapes[0];
        Assert.Equal(sourceWidth, copiedShape.Width, 1);
        Assert.Equal(sourceHeight, copiedShape.Height, 1);
    }

    [Fact]
    public void Reorder_ShouldChangeZOrder()
    {
        var pptPath = CreatePresentationWithTwoShapes("test_reorder.pptx");
        int shapeIndex;
        using (var pres = new Presentation(pptPath))
        {
            var nonPlaceholder = pres.Slides[0].Shapes.Where(s => s.Placeholder == null).ToList();
            shapeIndex = pres.Slides[0].Shapes.IndexOf(nonPlaceholder[0]);
        }

        var outputPath = CreateTestFilePath("test_reorder_output.pptx");
        var result = _tool.Execute("reorder", pptPath, slideIndex: 0, shapeIndex: shapeIndex, toIndex: 0,
            outputPath: outputPath);
        Assert.StartsWith("Shape Z-order changed", result);
    }

    [Fact]
    public void Align_ShouldAlignShapes()
    {
        var pptPath = CreatePresentationWithTwoShapes("test_align.pptx");
        int idx0, idx1;
        using (var pres = new Presentation(pptPath))
        {
            var sld = pres.Slides[0];
            var nonPlaceholder = sld.Shapes.Where(s => s.Placeholder == null).ToList();
            idx0 = sld.Shapes.IndexOf(nonPlaceholder[0]);
            idx1 = sld.Shapes.IndexOf(nonPlaceholder[1]);
        }

        var outputPath = CreateTestFilePath("test_align_output.pptx");
        var result = _tool.Execute("align", pptPath, slideIndex: 0, shapeIndices: [idx0, idx1], align: "left",
            outputPath: outputPath);
        Assert.StartsWith("Aligned", result);
    }

    [Fact]
    public void Flip_Horizontal_ShouldFlipShape()
    {
        var pptPath = CreateTestPresentation("test_flip_h.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_flip_h_output.pptx");
        var result = _tool.Execute("flip", pptPath, slideIndex: 0, shapeIndex: shapeIndex, flipHorizontal: true,
            outputPath: outputPath);
        Assert.StartsWith("Shape", result);
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
        Assert.StartsWith("Shape", result);
        Assert.Contains("flipped", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_get_{operation}.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0);
        Assert.Contains("\"totalCount\":", result);
    }

    [Theory]
    [InlineData("GET_DETAILS")]
    [InlineData("Get_Details")]
    [InlineData("get_details")]
    public void Operation_ShouldBeCaseInsensitive_GetDetails(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_details_{operation.Replace("_", "")}.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.Contains("\"position\":", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_del_{operation}.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath($"test_case_del_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex, outputPath: outputPath);
        Assert.StartsWith("Shape", result);
        Assert.Contains("deleted from slide", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_edit_{operation}.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex, x: 150,
            outputPath: outputPath);
        Assert.StartsWith("Shape", result);
        Assert.Contains("edited", result);
    }

    [Theory]
    [InlineData("SET_FORMAT")]
    [InlineData("Set_Format")]
    [InlineData("set_format")]
    public void Operation_ShouldBeCaseInsensitive_SetFormat(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_fmt_{operation.Replace("_", "")}.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath($"test_case_fmt_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex, fillColor: "#00FF00",
            outputPath: outputPath);
        Assert.StartsWith("Shape", result);
        Assert.Contains("format updated", result);
    }

    [Theory]
    [InlineData("CLEAR_FORMAT")]
    [InlineData("Clear_Format")]
    [InlineData("clear_format")]
    public void Operation_ShouldBeCaseInsensitive_ClearFormat(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_clr_{operation.Replace("_", "")}.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath($"test_case_clr_{operation.Replace("_", "")}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex, clearFill: true,
            outputPath: outputPath);
        Assert.StartsWith("Shape", result);
        Assert.Contains("format cleared", result);
    }

    [Theory]
    [InlineData("FLIP")]
    [InlineData("Flip")]
    [InlineData("flip")]
    public void Operation_ShouldBeCaseInsensitive_Flip(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_flip_{operation}.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath($"test_case_flip_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex, flipHorizontal: true,
            outputPath: outputPath);
        Assert.StartsWith("Shape", result);
        Assert.Contains("flipped", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Get_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_slide.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("get", pptPath, slideIndex: 99));
    }

    [Fact]
    public void Get_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_no_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", pptPath));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void GetDetails_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_no_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get_details", pptPath, slideIndex: 0));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_invalid_shape.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: 99, x: 100));
    }

    [Fact]
    public void Delete_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_del_no_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", pptPath, slideIndex: 0));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void ClearFormat_WithNoFlags_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_clear_no_flags.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("clear_format", pptPath, slideIndex: 0, shapeIndex: shapeIndex));
        Assert.Contains("At least one of clearFill or clearLine", ex.Message);
    }

    [Fact]
    public void Flip_WithNoDirection_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_flip_no_dir.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("flip", pptPath, slideIndex: 0, shapeIndex: shapeIndex));
        Assert.Contains("At least one of flipHorizontal or flipVertical", ex.Message);
    }

    [Fact]
    public void Group_WithLessThan2Shapes_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_group_one.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("group", pptPath, slideIndex: 0, shapeIndices: [shapeIndex]));
        Assert.Contains("At least 2 shapes", ex.Message);
    }

    [Fact]
    public void Ungroup_WithNonGroupShape_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_ungroup_non_group.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("ungroup", pptPath, slideIndex: 0, shapeIndex: shapeIndex));
        Assert.Contains("is not a group", ex.Message);
    }

    [Fact]
    public void Copy_WithoutFromSlide_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_copy_no_from.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("copy", pptPath, toSlide: 1, shapeIndex: 0));
        Assert.Contains("fromSlide is required", ex.Message);
    }

    [Fact]
    public void Copy_WithoutToSlide_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_copy_no_to.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("copy", pptPath, fromSlide: 0, shapeIndex: 0));
        Assert.Contains("toSlide is required", ex.Message);
    }

    [Fact]
    public void Reorder_WithoutToIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_reorder_no_to.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("reorder", pptPath, slideIndex: 0, shapeIndex: shapeIndex));
        Assert.Contains("toIndex is required", ex.Message);
    }

    [Fact]
    public void Align_WithoutAlign_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithTwoShapes("test_align_no_align.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("align", pptPath, slideIndex: 0, shapeIndices: [0, 1]));
        Assert.Contains("align is required", ex.Message);
    }

    [Fact]
    public void Align_WithLessThan2Shapes_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_align_one.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("align", pptPath, slideIndex: 0, shapeIndices: [0], align: "left"));
        Assert.Contains("At least 2 shapes", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void Get_WithSessionId_ShouldReturnShapesFromMemory()
    {
        var pptPath = CreateTestPresentation("test_session_get.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        Assert.Contains("\"userShapeIndex\":", result);
        Assert.Contains("\"totalCount\":", result);
    }

    [Fact]
    public void GetDetails_WithSessionId_ShouldReturnDetails()
    {
        var pptPath = CreateTestPresentation("test_session_details.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = FindNonPlaceholderShapeIndex(ppt.Slides[0]);
        var result = _tool.Execute("get_details", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.Contains("\"position\":", result);
        Assert.Contains("\"size\":", result);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_edit.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = FindNonPlaceholderShapeIndex(ppt.Slides[0]);
        var result = _tool.Execute("edit", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex, x: 300, y: 300);
        Assert.StartsWith("Shape", result);
        Assert.Contains("edited", result);
        Assert.Contains("session", result);
        if (!IsEvaluationMode())
        {
            var shape = ppt.Slides[0].Shapes[shapeIndex];
            Assert.Equal(300, shape.X, 1);
            Assert.Equal(300, shape.Y, 1);
        }
        else
        {
            // Fallback: verify basic structure in evaluation mode
            Assert.NotNull(ppt.Slides[0].Shapes[shapeIndex]);
        }
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.Count;
        var shapeIndex = FindNonPlaceholderShapeIndex(ppt.Slides[0]);
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex);
        Assert.StartsWith("Shape", result);
        Assert.Contains("deleted from slide", result);
        if (!IsEvaluationMode())
            Assert.True(ppt.Slides[0].Shapes.Count < initialCount);
        else
            // Fallback: verify basic structure in evaluation mode
            Assert.NotNull(ppt.Slides[0].Shapes);
    }

    [Fact]
    public void SetFormat_WithSessionId_ShouldModifyInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_format.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = FindNonPlaceholderShapeIndex(ppt.Slides[0]);
        var result = _tool.Execute("set_format", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            fillColor: "#00FF00");
        Assert.StartsWith("Shape", result);
        Assert.Contains("format updated", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session", slideIndex: 0));
    }

    [SkippableFact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode adds watermark shapes");
        var pptPath1 = CreateTestPresentation("test_path_shape.pptx");
        var pptPath2 = CreatePresentationWithTwoShapes("test_session_shape.pptx");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId, slideIndex: 0);
        Assert.Contains("\"userShapeCount\": 2", result);
    }

    #endregion
}