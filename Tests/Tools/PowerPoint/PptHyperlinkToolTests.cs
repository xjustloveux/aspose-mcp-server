using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Results.Common;
using AsposeMcpServer.Results.PowerPoint.Hyperlink;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

/// <summary>
///     Integration tests for PptHyperlinkTool.
///     Focuses on session management, file I/O, and operation routing.
///     Detailed parameter validation and business logic tests are in Handler tests.
/// </summary>
public class PptHyperlinkToolTests : PptTestBase
{
    private readonly PptHyperlinkTool _tool;

    public PptHyperlinkToolTests()
    {
        _tool = new PptHyperlinkTool(SessionManager);
    }

    private string CreatePresentationWithHyperlink(string fileName, string url = "https://example.com")
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        shape.HyperlinkClick = new Hyperlink(url);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private static int FindShapeIndex(string pptPath)
    {
        using var presentation = new Presentation(pptPath);
        var slide = presentation.Slides[0];
        var nonPlaceholderShapes = slide.Shapes.Where(s => s.Placeholder == null).ToList();
        if (nonPlaceholderShapes.Count == 0) nonPlaceholderShapes = slide.Shapes.ToList();
        foreach (var shape in nonPlaceholderShapes)
            if (Math.Abs(shape.X - 100) < 1 && Math.Abs(shape.Y - 100) < 1)
                return slide.Shapes.IndexOf(shape);
        return slide.Shapes.IndexOf(nonPlaceholderShapes[0]);
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Add_ShouldAddHyperlink()
    {
        var pptPath = CreatePresentationWithShape("test_add_hyperlink.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_hyperlink_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, shapeIndex: shapeIndex, url: "https://example.com",
            text: "Click here", outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Hyperlink added to slide", data.Message);
        using var presentation = new Presentation(outputPath);
        var shape = presentation.Slides[0].Shapes[shapeIndex];
        Assert.NotNull(shape.HyperlinkClick);
        Assert.Contains("example.com", shape.HyperlinkClick.ExternalUrl ?? "");
    }

    [Fact]
    public void Edit_ShouldModifyHyperlink()
    {
        var pptPath = CreatePresentationWithHyperlink("test_edit_hyperlink.pptx", "https://old.com");
        var outputPath = CreateTestFilePath("test_edit_hyperlink_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: 0, url: "https://new.com",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Hyperlink updated on slide", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Contains("new.com", presentation.Slides[0].Shapes[0].HyperlinkClick?.ExternalUrl ?? "");
    }

    [Fact]
    public void Delete_ShouldRemoveHyperlink()
    {
        var pptPath = CreatePresentationWithHyperlink("test_delete_hyperlink.pptx");
        var outputPath = CreateTestFilePath("test_delete_hyperlink_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Hyperlink deleted from slide", data.Message);
        using var presentation = new Presentation(outputPath);
        Assert.Null(presentation.Slides[0].Shapes[0].HyperlinkClick);
    }

    [Fact]
    public void Get_ShouldReturnAllHyperlinks()
    {
        var pptPath = CreatePresentationWithHyperlink("test_get_hyperlinks.pptx");
        var result = _tool.Execute("get", pptPath);
        var data = GetResultData<GetHyperlinksPptResult>(result);
        Assert.NotNull(data.TotalCount);
        Assert.NotNull(data.Slides);
    }

    #endregion

    #region Operation Routing

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive(string operation)
    {
        var pptPath = CreatePresentationWithShape($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, url: "https://example.com", text: "Link",
            outputPath: outputPath);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Hyperlink added to slide", data.Message);
    }

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithShape("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Get_WithSessionId_ShouldReturnHyperlinks()
    {
        var pptPath = CreatePresentationWithHyperlink("test_session_get_hyperlinks.pptx", "https://session-test.com");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        var data = GetResultData<GetHyperlinksPptResult>(result);
        Assert.NotNull(data.TotalCount);
        Assert.NotNull(data.Slides);
        var output = GetResultOutput<GetHyperlinksPptResult>(result);
        Assert.True(output.IsSession);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreatePresentationWithShape("test_session_add_hyperlink.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            url: "https://session-example.com", text: "Session Link");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Hyperlink added to slide", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        var shape = ppt.Slides[0].Shapes[shapeIndex];
        Assert.NotNull(shape.HyperlinkClick);
        Assert.Contains("session-example.com", shape.HyperlinkClick.ExternalUrl ?? "");
    }

    [Fact]
    public void Edit_WithSessionId_ShouldEditInMemory()
    {
        var pptPath = CreatePresentationWithHyperlink("test_session_edit_hyperlink.pptx", "https://old.com");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("edit", sessionId: sessionId, slideIndex: 0, shapeIndex: 0,
            url: "https://session-new.com");
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Hyperlink updated on slide", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        Assert.Contains("session-new.com", ppt.Slides[0].Shapes[0].HyperlinkClick?.ExternalUrl ?? "");
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithHyperlink("test_session_delete_hyperlink.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0, shapeIndex: 0);
        var data = GetResultData<SuccessResult>(result);
        Assert.StartsWith("Hyperlink deleted from slide", data.Message);
        var output = GetResultOutput<SuccessResult>(result);
        Assert.True(output.IsSession);
        Assert.Null(ppt.Slides[0].Shapes[0].HyperlinkClick);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        Assert.Throws<KeyNotFoundException>(() => _tool.Execute("get", sessionId: "invalid_session"));
    }

    [Fact]
    public void Execute_WithBothPathAndSessionId_ShouldPreferSessionId()
    {
        var pptPath1 = CreatePresentationWithShape("test_path_hyperlink.pptx");
        var pptPath2 = CreatePresentationWithHyperlink("test_session_hyperlink.pptx", "https://session-url.com");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId);
        var data = GetResultData<GetHyperlinksPptResult>(result);
        Assert.True(data.TotalCount > 0);
    }

    #endregion
}
