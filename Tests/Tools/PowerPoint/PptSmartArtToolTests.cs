using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptSmartArtTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptSmartArtToolTests : PptTestBase
{
    private readonly PptSmartArtTool _tool;

    public PptSmartArtToolTests()
    {
        _tool = new PptSmartArtTool(SessionManager);
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

    #region File I/O Smoke Tests

    [Fact]
    public void AddSmartArt_ShouldAddSmartArt()
    {
        var pptPath = CreatePresentation("test_add_smartart.pptx");
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

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, layout: "BasicProcess", outputPath: outputPath);
        Assert.StartsWith("SmartArt", result);
        Assert.Contains("added to slide", result);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown", pptPath, slideIndex: 0));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void AddSmartArt_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentation("test_session_add.pptx");
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
        var pptPath1 = CreatePresentation("test_path_smartart.pptx");
        var pptPath2 = CreatePresentation("test_session_smartart.pptx");

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
