using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptSmartArtToolTests : TestBase
{
    private readonly PptSmartArtTool _tool;

    public PptSmartArtToolTests()
    {
        _tool = new PptSmartArtTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private (string path, int shapeIndex) CreatePresentationWithSmartArt(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
        var shapeIndex = slide.Shapes.Count - 1;
        presentation.Save(filePath, SaveFormat.Pptx);
        return (filePath, shapeIndex);
    }

    #region General

    [Fact]
    public void AddSmartArt_ShouldAddSmartArt()
    {
        var pptPath = CreateTestPresentation("test_add_smartart.pptx");
        var outputPath = CreateTestFilePath("test_add_smartart_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, layout: "BasicProcess", x: 100, y: 100, width: 400,
            height: 300, outputPath: outputPath);
        Assert.StartsWith("SmartArt", result);
        Assert.Contains("added to slide", result);
        using var presentation = new Presentation(outputPath);
        var smartArts = presentation.Slides[0].Shapes.OfType<ISmartArt>().ToList();
        Assert.NotEmpty(smartArts);
    }

    [Fact]
    public void AddSmartArt_WithCustomPosition_ShouldPlaceAtPosition()
    {
        var pptPath = CreateTestPresentation("test_add_pos.pptx");
        var outputPath = CreateTestFilePath("test_add_pos_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, layout: "BasicCycle", x: 150, y: 200, width: 500, height: 400,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var smartArt = presentation.Slides[0].Shapes.OfType<ISmartArt>().First();
        Assert.Equal(150, smartArt.X, 1);
        Assert.Equal(200, smartArt.Y, 1);
    }

    [Theory]
    [InlineData("BasicProcess")]
    [InlineData("BasicCycle")]
    [InlineData("BasicPyramid")]
    [InlineData("Hierarchy")]
    public void AddSmartArt_WithVariousLayouts_ShouldSucceed(string layout)
    {
        var pptPath = CreateTestPresentation($"test_add_{layout}.pptx");
        var outputPath = CreateTestFilePath($"test_add_{layout}_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, layout: layout, outputPath: outputPath);
        Assert.StartsWith("SmartArt", result);
        Assert.Contains("added to slide", result);
    }

    [SkippableFact]
    public void ManageNodes_AddNode_ShouldAddNode()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_add_node.pptx");
        var outputPath = CreateTestFilePath("test_add_node_output.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex,
            action: "add", targetPath: targetPathJson, text: "New Node", outputPath: outputPath);
        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        using var ppt = new Presentation(outputPath);
        var smartArt = ppt.Slides[0].Shapes[shapeIndex] as ISmartArt;
        Assert.NotNull(smartArt);
        Assert.Contains(smartArt.AllNodes, n => n.TextFrame.Text.Contains("New Node"));
    }

    [Fact]
    public void ManageNodes_AddWithPosition_ShouldInsertAtPosition()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_add_position.pptx");
        var outputPath = CreateTestFilePath("test_add_position_output.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex,
            action: "add", targetPath: targetPathJson, text: "Inserted Node", position: 0, outputPath: outputPath);
        Assert.Contains("position 0", result);
        Assert.Contains("Inserted Node", result);
    }

    [SkippableFact]
    public void ManageNodes_EditNode_ShouldEditText()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_edit_node.pptx");
        var outputPath = CreateTestFilePath("test_edit_node_output.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex,
            action: "edit", targetPath: targetPathJson, text: "Edited Text", outputPath: outputPath);
        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
        using var ppt = new Presentation(outputPath);
        var smartArt = ppt.Slides[0].Shapes[shapeIndex] as ISmartArt;
        Assert.NotNull(smartArt);
        Assert.Contains("Edited Text", smartArt.AllNodes[0].TextFrame.Text);
    }

    [Fact]
    public void ManageNodes_DeleteNode_ShouldDeleteNode()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_delete_node.pptx");
        int initialNodeCount;
        using (var ppt = new Presentation(pptPath))
        {
            var smartArt = ppt.Slides[0].Shapes[shapeIndex] as ISmartArt;
            initialNodeCount = smartArt!.AllNodes.Count;
        }

        var outputPath = CreateTestFilePath("test_delete_node_output.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex,
            action: "delete", targetPath: targetPathJson, outputPath: outputPath);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        using var resultPpt = new Presentation(outputPath);
        var resultSmartArt = resultPpt.Slides[0].Shapes[shapeIndex] as ISmartArt;
        Assert.NotNull(resultSmartArt);
        Assert.True(resultSmartArt.AllNodes.Count < initialNodeCount);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, layout: "BasicProcess", outputPath: outputPath);
        Assert.StartsWith("SmartArt", result);
        Assert.Contains("added to slide", result);
    }

    [Theory]
    [InlineData("MANAGE_NODES")]
    [InlineData("Manage_Nodes")]
    [InlineData("manage_nodes")]
    public void Operation_ShouldBeCaseInsensitive_ManageNodes(string operation)
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt($"test_case_{operation.Replace("_", "")}.pptx");
        var outputPath = CreateTestFilePath($"test_case_{operation.Replace("_", "")}_output.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: shapeIndex,
            action: "delete", targetPath: targetPathJson, outputPath: outputPath);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void AddSmartArt_WithoutLayout_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_no_layout.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, slideIndex: 0));
        Assert.Contains("layout is required", ex.Message);
    }

    [Fact]
    public void AddSmartArt_WithInvalidLayout_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_invalid_layout.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, slideIndex: 0, layout: "InvalidLayoutName"));
        Assert.Contains("Invalid SmartArt layout", ex.Message);
    }

    [Fact]
    public void ManageNodes_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_manage_no_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage_nodes", pptPath, slideIndex: 0, action: "edit", targetPath: "[0]", text: "Test"));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void ManageNodes_WithoutAction_ShouldThrowArgumentException()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_manage_no_action.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex, targetPath: "[0]"));
        Assert.Contains("action is required", ex.Message);
    }

    [Fact]
    public void ManageNodes_WithoutTargetPath_ShouldThrowArgumentException()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_manage_no_target.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex, action: "edit",
                text: "Test"));
        Assert.Contains("targetPath is required", ex.Message);
    }

    [Fact]
    public void ManageNodes_WithNonSmartArtShape_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestFilePath("test_non_smartart.pptx");
        using (var pres = new Presentation())
        {
            pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            pres.Save(pptPath, SaveFormat.Pptx);
        }

        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: 0,
                action: "edit", targetPath: targetPathJson, text: "Test"));
        Assert.Contains("not a SmartArt", ex.Message);
    }

    [Fact]
    public void ManageNodes_WithInvalidRootIndex_ShouldThrowArgumentException()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_invalid_root.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 999 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex,
                action: "edit", targetPath: targetPathJson, text: "Test"));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void ManageNodes_AddWithInvalidPosition_ShouldThrowArgumentException()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_add_invalid_pos.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex,
                action: "add", targetPath: targetPathJson, text: "Test", position: 999));
        Assert.Contains("out of range", ex.Message);
    }

    [Fact]
    public void ManageNodes_AddWithoutText_ShouldThrowArgumentException()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_add_no_text.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex,
                action: "add", targetPath: targetPathJson));
        Assert.Contains("text parameter is required", ex.Message);
    }

    [Fact]
    public void ManageNodes_EditWithoutText_ShouldThrowArgumentException()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_edit_no_text.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex,
                action: "edit", targetPath: targetPathJson));
        Assert.Contains("text parameter is required", ex.Message);
    }

    [Fact]
    public void ManageNodes_WithUnknownAction_ShouldThrowArgumentException()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_unknown_action.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: shapeIndex,
                action: "invalid_action", targetPath: targetPathJson));
        Assert.Contains("Unknown action", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void AddSmartArt_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<ISmartArt>().Count();

        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, layout: "BasicProcess",
            x: 100, y: 100, width: 400, height: 300);
        Assert.StartsWith("SmartArt", result);
        Assert.Contains("added to slide", result);
        Assert.Contains("session", result);

        var smartArts = ppt.Slides[0].Shapes.OfType<ISmartArt>().Count();
        Assert.True(smartArts > initialCount);
    }

    [SkippableFact]
    public void ManageNodes_EditWithSessionId_ShouldEditInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_session_edit.pptx");
        var sessionId = OpenSession(pptPath);
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });

        var result = _tool.Execute("manage_nodes", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            action: "edit", targetPath: targetPathJson, text: "Session Edited");
        Assert.Contains("edited", result);
        Assert.Contains("session", result);

        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var smartArt = ppt.Slides[0].Shapes[shapeIndex] as ISmartArt;
        Assert.NotNull(smartArt);
        Assert.Contains("Session Edited", smartArt.AllNodes[0].TextFrame.Text);
    }

    [Fact]
    public void ManageNodes_AddNodeWithSessionId_ShouldAddInMemory()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_session_add_node.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var smartArt = ppt.Slides[0].Shapes[shapeIndex] as ISmartArt;
        var initialNodeCount = smartArt!.AllNodes.Count;

        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            action: "add", targetPath: targetPathJson, text: "New Session Node");
        Assert.Contains("added", result);
        Assert.Contains("session", result);
        Assert.True(smartArt.AllNodes.Count > initialNodeCount);
    }

    [Fact]
    public void ManageNodes_DeleteNodeWithSessionId_ShouldDeleteInMemory()
    {
        var (pptPath, shapeIndex) = CreatePresentationWithSmartArt("test_session_delete.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var smartArt = ppt.Slides[0].Shapes[shapeIndex] as ISmartArt;
        var initialNodeCount = smartArt!.AllNodes.Count;

        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            action: "delete", targetPath: targetPathJson);
        Assert.Contains("deleted", result);
        Assert.Contains("session", result);
        Assert.True(smartArt.AllNodes.Count < initialNodeCount);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute("add", sessionId: "invalid_session_id", slideIndex: 0, layout: "BasicProcess"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreateTestPresentation("test_path_smartart.pptx");
        var pptPath2 = CreateTestPresentation("test_session_smartart.pptx");

        var sessionId = OpenSession(pptPath2);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var initialCount = ppt.Slides[0].Shapes.OfType<ISmartArt>().Count();

        var result = _tool.Execute("add", pptPath1, sessionId, slideIndex: 0, layout: "BasicProcess");
        Assert.Contains("session", result);

        var smartArts = ppt.Slides[0].Shapes.OfType<ISmartArt>().Count();
        Assert.True(smartArts > initialCount);
    }

    #endregion
}