using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptSlideToolTests : TestBase
{
    private readonly PptSlideTool _tool;

    public PptSlideToolTests()
    {
        _tool = new PptSlideTool(SessionManager);
    }

    private string CreatePptPresentation(string fileName, int slideCount = 2)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePptWithShape(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General

    [Fact]
    public void Add_ShouldAddNewSlide()
    {
        var pptPath = CreatePptPresentation("test_add.pptx");
        var outputPath = CreateTestFilePath("test_add_output.pptx");
        var result = _tool.Execute("add", pptPath, outputPath: outputPath);
        Assert.StartsWith("Slide added", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(3, presentation.Slides.Count);
    }

    [Fact]
    public void Add_WithLayoutType_ShouldUseLayout()
    {
        var pptPath = CreatePptPresentation("test_add_layout.pptx");
        var outputPath = CreateTestFilePath("test_add_layout_output.pptx");
        var result = _tool.Execute("add", pptPath, layoutType: "Title", outputPath: outputPath);
        Assert.StartsWith("Slide added", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(3, presentation.Slides.Count);
    }

    [Fact]
    public void Delete_ShouldDeleteSlide()
    {
        var pptPath = CreatePptPresentation("test_delete.pptx", 3);
        var outputPath = CreateTestFilePath("test_delete_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 1, outputPath: outputPath);
        Assert.StartsWith("Slide 1 deleted", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(2, presentation.Slides.Count);
    }

    [Fact]
    public void GetInfo_ShouldReturnSlidesInfo()
    {
        var pptPath = CreatePptPresentation("test_get_info.pptx", 3);
        var result = _tool.Execute("get_info", pptPath);
        Assert.Contains("\"count\": 3", result);
        Assert.Contains("\"layoutType\":", result);
        Assert.Contains("\"layoutName\":", result);
        Assert.Contains("\"availableLayouts\":", result);
    }

    [Fact]
    public void Move_ShouldMoveSlide()
    {
        var pptPath = CreatePptPresentation("test_move.pptx", 3);
        var outputPath = CreateTestFilePath("test_move_output.pptx");
        var result = _tool.Execute("move", pptPath, fromIndex: 2, toIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Slide moved", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(3, presentation.Slides.Count);
    }

    [Fact]
    public void Duplicate_ShouldDuplicateSlide()
    {
        var pptPath = CreatePptPresentation("test_duplicate.pptx");
        var outputPath = CreateTestFilePath("test_duplicate_output.pptx");
        var result = _tool.Execute("duplicate", pptPath, slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Slide 0 duplicated", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(3, presentation.Slides.Count);
    }

    [Fact]
    public void Duplicate_WithInsertAt_ShouldInsertAtPosition()
    {
        var pptPath = CreatePptPresentation("test_duplicate_insert.pptx", 3);
        var outputPath = CreateTestFilePath("test_duplicate_insert_output.pptx");
        var result = _tool.Execute("duplicate", pptPath, slideIndex: 0, insertAt: 2, outputPath: outputPath);
        Assert.StartsWith("Slide 0 duplicated", result);
        using var presentation = new Presentation(outputPath);
        Assert.Equal(4, presentation.Slides.Count);
    }

    [Fact]
    public void Hide_ShouldHideSlide()
    {
        var pptPath = CreatePptPresentation("test_hide.pptx");
        var outputPath = CreateTestFilePath("test_hide_output.pptx");
        var result = _tool.Execute("hide", pptPath, slideIndex: 0, hidden: true, outputPath: outputPath);
        Assert.StartsWith("Set", result);
        Assert.Contains("hidden=True", result);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        if (!IsEvaluationMode())
            Assert.True(presentation.Slides[0].Hidden);
        else
            Assert.True(presentation.Slides.Count > 0, "Fallback: presentation should have slides");
    }

    [Fact]
    public void Hide_WithMultipleSlides_ShouldHideAll()
    {
        var pptPath = CreatePptPresentation("test_hide_multi.pptx", 3);
        var outputPath = CreateTestFilePath("test_hide_multi_output.pptx");
        var result = _tool.Execute("hide", pptPath, slideIndices: "[0,1]", hidden: true, outputPath: outputPath);
        Assert.Contains("2 slide(s) hidden=True", result);
        Assert.True(File.Exists(outputPath));
        using var presentation = new Presentation(outputPath);
        Assert.Equal(3, presentation.Slides.Count);
        if (!IsEvaluationMode())
        {
            Assert.True(presentation.Slides[0].Hidden);
            Assert.True(presentation.Slides[1].Hidden);
            Assert.False(presentation.Slides[2].Hidden);
        }
        else
        {
            // Fallback: verify basic structure in evaluation mode
            Assert.NotNull(presentation.Slides[0]);
            Assert.NotNull(presentation.Slides[1]);
            Assert.NotNull(presentation.Slides[2]);
        }
    }

    [SkippableFact]
    public void Clear_ShouldClearSlideContent()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Clear verification may not work correctly in evaluation mode");
        var pptPath = CreatePptWithShape("test_clear.pptx");
        var outputPath = CreateTestFilePath("test_clear_output.pptx");
        var result = _tool.Execute("clear", pptPath, slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Cleared all shapes", result);
        using var presentation = new Presentation(outputPath);
        Assert.Empty(presentation.Slides[0].Shapes);
    }

    [Fact]
    public void Edit_ShouldEditSlide()
    {
        var pptPath = CreatePptPresentation("test_edit.pptx");
        var outputPath = CreateTestFilePath("test_edit_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Slide 0 updated", result);
    }

    [Fact]
    public void Edit_WithLayoutIndex_ShouldChangeLayout()
    {
        var pptPath = CreatePptPresentation("test_edit_layout.pptx");
        var outputPath = CreateTestFilePath("test_edit_layout_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, layoutIndex: 1, outputPath: outputPath);
        Assert.StartsWith("Slide 0 updated", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pptPath = CreatePptPresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, outputPath: outputPath);
        Assert.StartsWith("Slide added", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var pptPath = CreatePptPresentation($"test_case_del_{operation}.pptx", 3);
        var outputPath = CreateTestFilePath($"test_case_del_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 1, outputPath: outputPath);
        Assert.StartsWith("Slide 1 deleted", result);
    }

    [Theory]
    [InlineData("GET_INFO")]
    [InlineData("Get_Info")]
    [InlineData("get_info")]
    public void Operation_ShouldBeCaseInsensitive_GetInfo(string operation)
    {
        var pptPath = CreatePptPresentation($"test_case_info_{operation.Replace("_", "")}.pptx");
        var result = _tool.Execute(operation, pptPath);
        Assert.Contains("\"count\":", result);
    }

    [Theory]
    [InlineData("MOVE")]
    [InlineData("Move")]
    [InlineData("move")]
    public void Operation_ShouldBeCaseInsensitive_Move(string operation)
    {
        var pptPath = CreatePptPresentation($"test_case_move_{operation}.pptx", 3);
        var outputPath = CreateTestFilePath($"test_case_move_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, fromIndex: 2, toIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Slide moved", result);
    }

    [Theory]
    [InlineData("DUPLICATE")]
    [InlineData("Duplicate")]
    [InlineData("duplicate")]
    public void Operation_ShouldBeCaseInsensitive_Duplicate(string operation)
    {
        var pptPath = CreatePptPresentation($"test_case_dup_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_dup_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Slide 0 duplicated", result);
    }

    [Theory]
    [InlineData("HIDE")]
    [InlineData("Hide")]
    [InlineData("hide")]
    public void Operation_ShouldBeCaseInsensitive_Hide(string operation)
    {
        var pptPath = CreatePptPresentation($"test_case_hide_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_hide_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, hidden: true, outputPath: outputPath);
        Assert.StartsWith("Set", result);
        Assert.Contains("hidden=True", result);
    }

    [Theory]
    [InlineData("CLEAR")]
    [InlineData("Clear")]
    [InlineData("clear")]
    public void Operation_ShouldBeCaseInsensitive_Clear(string operation)
    {
        var pptPath = CreatePptWithShape($"test_case_clear_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_clear_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Cleared", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var pptPath = CreatePptPresentation($"test_case_edit_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Slide 0 updated", result);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Delete_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_delete_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", pptPath));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void Delete_WithInvalidIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_delete_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", pptPath, slideIndex: 99));
        Assert.Contains("slideIndex must be between", ex.Message);
    }

    [Fact]
    public void Delete_LastSlide_ShouldThrowInvalidOperationException()
    {
        var pptPath = CreateTestFilePath("test_delete_last.pptx");
        using (var ppt = new Presentation())
        {
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var ex = Assert.Throws<InvalidOperationException>(() => _tool.Execute("delete", pptPath, slideIndex: 0));
        Assert.Contains("Cannot delete the last slide", ex.Message);
    }

    [Fact]
    public void Move_WithoutFromIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_move_no_from.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("move", pptPath, toIndex: 0));
        Assert.Contains("fromIndex is required", ex.Message);
    }

    [Fact]
    public void Move_WithoutToIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_move_no_to.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("move", pptPath, fromIndex: 0));
        Assert.Contains("toIndex is required", ex.Message);
    }

    [Fact]
    public void Move_WithInvalidFromIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_move_invalid_from.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("move", pptPath, fromIndex: 99, toIndex: 0));
        Assert.Contains("fromIndex must be between", ex.Message);
    }

    [Fact]
    public void Move_WithInvalidToIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_move_invalid_to.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("move", pptPath, fromIndex: 0, toIndex: 99));
        Assert.Contains("toIndex must be between", ex.Message);
    }

    [Fact]
    public void Duplicate_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_dup_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("duplicate", pptPath));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void Duplicate_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_dup_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("duplicate", pptPath, slideIndex: 99));
        Assert.Contains("slideIndex must be between", ex.Message);
    }

    [Fact]
    public void Duplicate_WithInvalidInsertAt_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_dup_invalid_insert.pptx");
        var ex =
            Assert.Throws<ArgumentException>(() => _tool.Execute("duplicate", pptPath, slideIndex: 0, insertAt: 99));
        Assert.Contains("insertAt must be between", ex.Message);
    }

    [Fact]
    public void Hide_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_hide_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("hide", pptPath, slideIndices: "[99]", hidden: true));
        Assert.Contains("slide index 99 out of range", ex.Message);
    }

    [Fact]
    public void Clear_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_clear_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("clear", pptPath));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void Clear_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_clear_invalid.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("clear", pptPath, slideIndex: 99));
        Assert.Contains("Slide index", ex.Message);
    }

    [Fact]
    public void Edit_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_edit_no_index.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("edit", pptPath));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void Edit_WithInvalidLayoutIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePptPresentation("test_edit_invalid_layout.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("edit", pptPath, slideIndex: 0, layoutIndex: 99));
        Assert.Contains("layoutIndex must be between", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePptPresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides.Count;
        var result = _tool.Execute("add", sessionId: sessionId);
        Assert.StartsWith("Slide added", result);
        Assert.Contains("session", result);
        Assert.Equal(initialCount + 1, ppt.Slides.Count);
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePptPresentation("test_session_delete.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 1);
        Assert.StartsWith("Slide 1 deleted", result);
        Assert.Equal(2, ppt.Slides.Count);
    }

    [Fact]
    public void GetInfo_WithSessionId_ShouldReturnInfo()
    {
        var pptPath = CreatePptPresentation("test_session_info.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get_info", sessionId: sessionId);
        Assert.Contains("\"count\":", result);
        Assert.Contains("\"layoutType\":", result);
    }

    [Fact]
    public void Move_WithSessionId_ShouldMoveInMemory()
    {
        var pptPath = CreatePptPresentation("test_session_move.pptx", 3);
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("move", sessionId: sessionId, fromIndex: 2, toIndex: 0);
        Assert.StartsWith("Slide moved", result);
    }

    [Fact]
    public void Duplicate_WithSessionId_ShouldDuplicateInMemory()
    {
        var pptPath = CreatePptPresentation("test_session_dup.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides.Count;
        var result = _tool.Execute("duplicate", sessionId: sessionId, slideIndex: 0);
        Assert.StartsWith("Slide 0 duplicated", result);
        Assert.Equal(initialCount + 1, ppt.Slides.Count);
    }

    [Fact]
    public void Hide_WithSessionId_ShouldHideInMemory()
    {
        var pptPath = CreatePptPresentation("test_session_hide.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("hide", sessionId: sessionId, slideIndex: 0, hidden: true);
        Assert.StartsWith("Set", result);
        Assert.Contains("hidden=True", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Clear_WithSessionId_ShouldClearInMemory()
    {
        var pptPath = CreatePptWithShape("test_session_clear.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("clear", sessionId: sessionId, slideIndex: 0);
        Assert.StartsWith("Cleared", result);
        Assert.Empty(ppt.Slides[0].Shapes);
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var pptPath = CreatePptPresentation("test_session_edit.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("edit", sessionId: sessionId, slideIndex: 0);
        Assert.StartsWith("Slide 0 updated", result);
        Assert.Contains("session", result);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get_info", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePptPresentation("test_path_ppt.pptx");
        var pptPath2 = CreatePptPresentation("test_session_ppt.pptx", 5);
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get_info", pptPath1, sessionId);
        Assert.Contains("\"count\": 5", result);
    }

    #endregion
}