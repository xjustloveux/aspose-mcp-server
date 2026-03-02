using System.Runtime.Versioning;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Shape;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptShapeTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
[SupportedOSPlatform("windows")]
public class PptShapeToolTests : PptTestBase
{
    private readonly PptShapeTool _tool;

    public PptShapeToolTests()
    {
        _tool = new PptShapeTool(SessionManager);
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

    #region File I/O Smoke Tests

    [SkippableFact]
    public void Get_ShouldReturnAllShapes()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_get.pptx");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        var data = GetResultData<GetShapesResult>(result);
        Assert.True(data.Count >= 0);
        Assert.NotNull(data.Shapes);
    }

    [SkippableFact]
    public void GetDetails_ShouldReturnShapeDetails()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_get_details.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var result = _tool.Execute("get_details", pptPath, slideIndex: 0, shapeIndex: shapeIndex);
        var data = GetResultData<GetShapeDetailsResult>(result);
        Assert.Equal(shapeIndex, data.Index);
        Assert.True(data.X >= 0);
    }

    [SkippableFact]
    public void Delete_ShouldRemoveShape()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_delete.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_delete_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: shapeIndex, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Shape", data.Message);
        Assert.Contains("deleted from slide", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Edit_ShouldModifyPosition()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_edit_pos.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_edit_pos_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: shapeIndex, x: 200, y: 200,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Shape", data.Message);
        Assert.Contains("updated", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void SetFormat_ShouldSetFillColor()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_set_format.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_set_format_output.pptx");
        var result = _tool.Execute("set_format", pptPath, slideIndex: 0, shapeIndex: shapeIndex, fillColor: "#FF0000",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Format applied to shape", data.Message);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void ClearFormat_ShouldClearFill()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_clear_format.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_clear_format_output.pptx");
        var result = _tool.Execute("clear_format", pptPath, slideIndex: 0, shapeIndex: shapeIndex, clearFill: true,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Format cleared from shape", data.Message);
    }

    [SkippableFact]
    public void Group_ShouldGroupShapes()
    {
        SkipIfNotWindows();
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
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Grouped", data.Message);
    }

    [SkippableFact]
    public void Flip_ShouldFlipShape()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_flip.pptx");
        var shapeIndex = FindNonPlaceholderShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_flip_output.pptx");
        var result = _tool.Execute("flip", pptPath, slideIndex: 0, shapeIndex: shapeIndex, flipHorizontal: true,
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Shape", data.Message);
        Assert.Contains("flipped", data.Message);
    }

    #endregion

    #region Operation Routing

    [SkippableTheory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape($"test_case_{operation}.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0);
        var data = GetResultData<GetShapesResult>(result);
        Assert.True(data.Count >= 0);
    }

    [SkippableFact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [SkippableFact]
    public void Execute_WithNoPathOrSessionId_ShouldThrowException()
    {
        SkipIfNotWindows();
        Assert.ThrowsAny<Exception>(() => _tool.Execute("get", slideIndex: 0));
    }

    #endregion

    #region Session Management

    [SkippableFact]
    public void Get_WithSessionId_ShouldReturnShapesFromMemory()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_session_get.pptx");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId, slideIndex: 0);
        var data = GetResultData<GetShapesResult>(result);
        Assert.True(data.Count >= 0);
        Assert.NotNull(data.Shapes);
    }

    [SkippableFact]
    public void GetDetails_WithSessionId_ShouldReturnDetails()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_session_details.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = FindNonPlaceholderShapeIndex(ppt.Slides[0]);
        var result = _tool.Execute("get_details", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex);
        var data = GetResultData<GetShapeDetailsResult>(result);
        Assert.Equal(shapeIndex, data.Index);
        Assert.True(data.X >= 0);
    }

    [SkippableFact]
    public void Edit_WithSessionId_ShouldModifyInMemory()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_session_edit.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = FindNonPlaceholderShapeIndex(ppt.Slides[0]);
        var result = _tool.Execute("edit", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex, x: 300, y: 300);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Shape", data.Message);
        Assert.Contains("updated", data.Message);
    }

    [SkippableFact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = FindNonPlaceholderShapeIndex(ppt.Slides[0]);
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Shape", data.Message);
        Assert.Contains("deleted from slide", data.Message);
    }

    [SkippableFact]
    public void SetFormat_WithSessionId_ShouldModifyInMemory()
    {
        SkipIfNotWindows();
        var pptPath = CreatePresentationWithShape("test_session_format.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var shapeIndex = FindNonPlaceholderShapeIndex(ppt.Slides[0]);
        var result = _tool.Execute("set_format", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            fillColor: "#00FF00");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Format applied to shape", data.Message);
    }

    [SkippableFact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        SkipIfNotWindows();
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session", slideIndex: 0));
    }

    [SkippableFact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        SkipIfNotWindows();
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode adds watermark shapes");
        var pptPath1 = CreatePresentationWithShape("test_path_shape.pptx");
        var pptPath2 = CreatePresentationWithTwoShapes("test_session_shape.pptx");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId, slideIndex: 0);
        var data = GetResultData<GetShapesResult>(result);
        Assert.True(data.Count >= 0);
        Assert.NotNull(data.Shapes);
    }

    #endregion
}
