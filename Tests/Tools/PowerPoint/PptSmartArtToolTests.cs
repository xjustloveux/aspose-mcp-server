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

    #region General Tests

    [Fact]
    public void AddSmartArt_ShouldAddSmartArt()
    {
        var pptPath = CreateTestPresentation("test_add_smartart.pptx");
        var outputPath = CreateTestFilePath("test_add_smartart_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, layout: "BasicProcess", x: 100, y: 100, width: 400, height: 300,
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var smartArts = slide.Shapes.OfType<ISmartArt>().ToList();
        Assert.True(smartArts.Count > 0, "Slide should contain at least one SmartArt");
    }

    [SkippableFact]
    public void GetSmartArt_ShouldReturnSmartArtInfo()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestPresentation("test_manage_smartart_nodes.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            var smartArt = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            presentation.Save(pptPath, SaveFormat.Pptx);

            smartArtShapeIndex = -1;
            for (var i = 0; i < slide.Shapes.Count; i++)
                if (slide.Shapes[i] == smartArt)
                {
                    smartArtShapeIndex = i;
                    break;
                }

            Assert.True(smartArtShapeIndex >= 0, "SmartArt shape should be found in slide");
        }

        var outputPath = CreateTestFilePath("test_manage_smartart_nodes_output.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: smartArtShapeIndex,
            action: "add", targetPath: targetPathJson, text: "New Node", outputPath: outputPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("SmartArt", result, StringComparison.OrdinalIgnoreCase);
        using var ppt = new Presentation(outputPath);
        var smartArt2 = ppt.Slides[0].Shapes[smartArtShapeIndex] as ISmartArt;
        Assert.NotNull(smartArt2);
        Assert.True(smartArt2.AllNodes.Any(n => n.TextFrame.Text.Contains("New Node")),
            "Node should contain 'New Node'");
    }

    [Fact]
    public void AddSmartArt_InvalidLayout_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_invalid_layout.pptx");
        Assert.Throws<ArgumentException>(() =>
            _tool.Execute("add", pptPath, slideIndex: 0, layout: "InvalidLayoutName"));
    }

    [Fact]
    public void ManageNodes_InvalidRootIndex_ShouldThrow()
    {
        var pptPath = CreateTestFilePath("test_invalid_root_index.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            _ = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var targetPathJson = JsonSerializer.Serialize(new[] { 999 });
        Assert.Throws<ArgumentException>(() => _tool.Execute("manage_nodes", pptPath, slideIndex: 0,
            shapeIndex: smartArtShapeIndex, action: "edit", targetPath: targetPathJson, text: "Test"));
    }

    [SkippableFact]
    public void ManageNodes_EditNode_ShouldEditText()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation mode truncates text content");

        var pptPath = CreateTestFilePath("test_edit_node.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            _ = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_edit_node_output.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: smartArtShapeIndex,
            action: "edit", targetPath: targetPathJson, text: "Edited Text", outputPath: outputPath);
        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
        using var ppt = new Presentation(outputPath);
        var smartArt = ppt.Slides[0].Shapes[smartArtShapeIndex] as ISmartArt;
        Assert.NotNull(smartArt);
        Assert.Contains("Edited Text", smartArt.AllNodes[0].TextFrame.Text);
    }

    [Fact]
    public void ManageNodes_DeleteRootNode_ShouldSucceed()
    {
        var pptPath = CreateTestFilePath("test_delete_root_node.pptx");
        int smartArtShapeIndex;
        int initialNodeCount;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var smartArt = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            initialNodeCount = smartArt.AllNodes.Count;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_root_node_output.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: smartArtShapeIndex,
            action: "delete", targetPath: targetPathJson, outputPath: outputPath);
        Assert.Contains("deleted", result, StringComparison.OrdinalIgnoreCase);
        using var resultPres = new Presentation(outputPath);
        var resultSmartArt = resultPres.Slides[0].Shapes[smartArtShapeIndex] as ISmartArt;
        Assert.NotNull(resultSmartArt);
        Assert.True(resultSmartArt.AllNodes.Count < initialNodeCount);
    }

    [Fact]
    public void ManageNodes_NotSmartArtShape_ShouldThrow()
    {
        var pptPath = CreateTestFilePath("test_not_smartart.pptx");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        Assert.Throws<ArgumentException>(() => _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: 0,
            action: "edit", targetPath: targetPathJson, text: "Test"));
    }

    [Fact]
    public void ManageNodes_AddWithPosition_ShouldInsertAtPosition()
    {
        var pptPath = CreateTestFilePath("test_add_with_position.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            _ = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_add_with_position_output.pptx");
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", pptPath, slideIndex: 0, shapeIndex: smartArtShapeIndex,
            action: "add", targetPath: targetPathJson, text: "Inserted Node", position: 0, outputPath: outputPath);
        Assert.Contains("position 0", result);
        Assert.Contains("Inserted Node", result);
    }

    [Fact]
    public void ManageNodes_AddWithInvalidPosition_ShouldThrow()
    {
        var pptPath = CreateTestFilePath("test_add_invalid_position.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            _ = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        Assert.Throws<ArgumentException>(() => _tool.Execute("manage_nodes", pptPath, slideIndex: 0,
            shapeIndex: smartArtShapeIndex, action: "add", targetPath: targetPathJson, text: "Test", position: 999));
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath, slideIndex: 0));
    }

    [Fact]
    public void Add_MissingLayout_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_missing_layout.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("add", pptPath, slideIndex: 0));
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void AddSmartArt_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add_smartart.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var slide = ppt.Slides[0];
        var initialShapeCount = slide.Shapes.Count;
        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, layout: "BasicProcess",
            x: 100, y: 100, width: 400, height: 300);
        Assert.Contains("SmartArt", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("session", result);
        Assert.True(slide.Shapes.Count > initialShapeCount);
        var smartArts = slide.Shapes.OfType<ISmartArt>().ToList();
        Assert.True(smartArts.Count > 0, "Slide should contain SmartArt in memory");
    }

    [SkippableFact]
    public void ManageNodes_WithSessionId_ShouldEditInMemory()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides,
            "Evaluation mode adds watermarks that interfere with text assertions");

        var pptPath = CreateTestFilePath("test_session_manage_nodes.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            _ = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slide.Shapes.Count - 1;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", sessionId: sessionId, slideIndex: 0,
            shapeIndex: smartArtShapeIndex, action: "edit", targetPath: targetPathJson, text: "Session Edited");
        Assert.Contains("edited", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var smartArt = ppt.Slides[0].Shapes[smartArtShapeIndex] as ISmartArt;
        Assert.NotNull(smartArt);
        var nodeText = smartArt.AllNodes[0].TextFrame.Text;
        Assert.Contains("Session Edited", nodeText);
    }

    [Fact]
    public void ManageNodes_AddNode_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestFilePath("test_session_add_node.pptx");
        int smartArtShapeIndex;
        int initialNodeCount;
        using (var presentation = new Presentation())
        {
            var slideToSetup = presentation.Slides[0];
            var smartArtToSetup = slideToSetup.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            smartArtShapeIndex = slideToSetup.Shapes.Count - 1;
            initialNodeCount = smartArtToSetup.AllNodes.Count;
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var targetPathJson = JsonSerializer.Serialize(new[] { 0 });
        var result = _tool.Execute("manage_nodes", sessionId: sessionId, slideIndex: 0,
            shapeIndex: smartArtShapeIndex, action: "add", targetPath: targetPathJson, text: "New Session Node");
        Assert.Contains("added", result, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var smartArt = ppt.Slides[0].Shapes[smartArtShapeIndex] as ISmartArt;
        Assert.NotNull(smartArt);
        Assert.True(smartArt.AllNodes.Count > initialNodeCount);
    }

    #endregion
}