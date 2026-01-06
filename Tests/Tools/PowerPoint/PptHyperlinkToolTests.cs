using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.Tools.PowerPoint;

public class PptHyperlinkToolTests : TestBase
{
    private readonly PptHyperlinkTool _tool;

    public PptHyperlinkToolTests()
    {
        _tool = new PptHyperlinkTool(SessionManager);
    }

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
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

    private string CreatePresentationWithMultipleSlides(string fileName, int slideCount = 3)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        for (var i = 1; i < slideCount; i++)
            presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    private string CreatePresentationWithPortionHyperlink(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 50);
        shape.TextFrame.Paragraphs.Clear();
        var paragraph = new Paragraph();
        paragraph.Portions.Add(new Portion("Click "));
        var linkPortion = new Portion("here")
        {
            PortionFormat = { HyperlinkClick = new Hyperlink("https://portion-link.com") }
        };
        paragraph.Portions.Add(linkPortion);
        paragraph.Portions.Add(new Portion(" for more"));
        shape.TextFrame.Paragraphs.Add(paragraph);
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

    #region General

    [Fact]
    public void Add_ShouldAddHyperlink()
    {
        var pptPath = CreateTestPresentation("test_add_hyperlink.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var outputPath = CreateTestFilePath("test_add_hyperlink_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, shapeIndex: shapeIndex, url: "https://example.com",
            text: "Click here", outputPath: outputPath);
        Assert.StartsWith("Hyperlink added to slide", result);
        using var presentation = new Presentation(outputPath);
        var shape = presentation.Slides[0].Shapes[shapeIndex];
        Assert.NotNull(shape.HyperlinkClick);
        Assert.Contains("example.com", shape.HyperlinkClick.ExternalUrl ?? "");
    }

    [Fact]
    public void Add_WithNewShape_ShouldCreateShapeWithLink()
    {
        var pptPath = CreateTestPresentation("test_add_new_shape.pptx");
        var outputPath = CreateTestFilePath("test_add_new_shape_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, url: "https://example.com", text: "New Link", x: 100,
            y: 200, width: 150, height: 40, outputPath: outputPath);
        Assert.StartsWith("Hyperlink added to slide", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Add_WithSlideTargetIndex_ShouldAddInternalLink()
    {
        var pptPath = CreatePresentationWithMultipleSlides("test_add_internal_link.pptx");
        var outputPath = CreateTestFilePath("test_add_internal_link_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, slideTargetIndex: 1, text: "Go to slide 2",
            outputPath: outputPath);
        Assert.StartsWith("Hyperlink added to slide", result);
        Assert.True(File.Exists(outputPath));
    }

    [SkippableFact]
    public void Add_WithLinkText_ShouldAddPortionLevelHyperlink()
    {
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation watermark interferes with text content");
        var pptPath = CreateTestPresentation("test_add_portion_hyperlink.pptx");
        var outputPath = CreateTestFilePath("test_add_portion_hyperlink_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, url: "https://example.com",
            text: "Please click here for more info", linkText: "here", outputPath: outputPath);
        Assert.StartsWith("Hyperlink added to slide", result);
        using var presentation = new Presentation(outputPath);
        var autoShape = presentation.Slides[0].Shapes.OfType<IAutoShape>().Last();
        var herePortion = autoShape.TextFrame.Paragraphs.SelectMany(p => p.Portions)
            .FirstOrDefault(p => p.Text == "here");
        Assert.NotNull(herePortion);
        Assert.NotNull(herePortion.PortionFormat.HyperlinkClick);
    }

    [Fact]
    public void Edit_ShouldModifyHyperlink()
    {
        var pptPath = CreatePresentationWithHyperlink("test_edit_hyperlink.pptx", "https://old.com");
        var outputPath = CreateTestFilePath("test_edit_hyperlink_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: 0, url: "https://new.com",
            outputPath: outputPath);
        Assert.StartsWith("Hyperlink updated on slide", result);
        using var presentation = new Presentation(outputPath);
        Assert.Contains("new.com", presentation.Slides[0].Shapes[0].HyperlinkClick?.ExternalUrl ?? "");
    }

    [Fact]
    public void Edit_WithRemoveHyperlink_ShouldRemoveLink()
    {
        var pptPath = CreatePresentationWithHyperlink("test_edit_remove.pptx");
        var outputPath = CreateTestFilePath("test_edit_remove_output.pptx");
        var result = _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: 0, removeHyperlink: true,
            outputPath: outputPath);
        Assert.StartsWith("Hyperlink updated on slide", result);
        using var presentation = new Presentation(outputPath);
        Assert.Null(presentation.Slides[0].Shapes[0].HyperlinkClick);
    }

    [Fact]
    public void Delete_ShouldRemoveHyperlink()
    {
        var pptPath = CreatePresentationWithHyperlink("test_delete_hyperlink.pptx");
        var outputPath = CreateTestFilePath("test_delete_hyperlink_output.pptx");
        var result = _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Hyperlink deleted from slide", result);
        using var presentation = new Presentation(outputPath);
        Assert.Null(presentation.Slides[0].Shapes[0].HyperlinkClick);
    }

    [Fact]
    public void Delete_ShouldRemovePortionLevelHyperlinks()
    {
        var pptPath = CreatePresentationWithPortionHyperlink("test_delete_portion_hyperlink.pptx");
        var outputPath = CreateTestFilePath("test_delete_portion_hyperlink_output.pptx");
        _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var autoShape = presentation.Slides[0].Shapes.OfType<IAutoShape>().First();
        foreach (var para in autoShape.TextFrame.Paragraphs)
        foreach (var portion in para.Portions)
            Assert.Null(portion.PortionFormat.HyperlinkClick);
    }

    [Fact]
    public void Get_ShouldReturnAllHyperlinks()
    {
        var pptPath = CreatePresentationWithHyperlink("test_get_hyperlinks.pptx");
        var result = _tool.Execute("get", pptPath);
        Assert.NotNull(result);
        Assert.Contains("Hyperlink", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Get_WithSlideIndex_ShouldReturnSlideHyperlinks()
    {
        var pptPath = CreatePresentationWithHyperlink("test_get_slide_hyperlinks.pptx");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        Assert.Contains("slideIndex", result);
        Assert.Contains("hyperlinks", result);
    }

    [Fact]
    public void Get_ShouldReturnPortionLevelHyperlinks()
    {
        var pptPath = CreatePresentationWithPortionHyperlink("test_get_portion_hyperlinks.pptx");
        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        Assert.Contains("\"level\": \"text\"", result);
        Assert.Contains("\"text\": \"here\"", result);
        Assert.Contains("portion-link.com", result);
    }

    [Theory]
    [InlineData("ADD")]
    [InlineData("Add")]
    [InlineData("add")]
    public void Operation_ShouldBeCaseInsensitive_Add(string operation)
    {
        var pptPath = CreateTestPresentation($"test_case_add_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_add_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, url: "https://example.com", text: "Link",
            outputPath: outputPath);
        Assert.StartsWith("Hyperlink added to slide", result);
    }

    [Theory]
    [InlineData("EDIT")]
    [InlineData("Edit")]
    [InlineData("edit")]
    public void Operation_ShouldBeCaseInsensitive_Edit(string operation)
    {
        var pptPath = CreatePresentationWithHyperlink($"test_case_edit_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_edit_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: 0, url: "https://new.com",
            outputPath: outputPath);
        Assert.StartsWith("Hyperlink updated on slide", result);
    }

    [Theory]
    [InlineData("DELETE")]
    [InlineData("Delete")]
    [InlineData("delete")]
    public void Operation_ShouldBeCaseInsensitive_Delete(string operation)
    {
        var pptPath = CreatePresentationWithHyperlink($"test_case_delete_{operation}.pptx");
        var outputPath = CreateTestFilePath($"test_case_delete_{operation}_output.pptx");
        var result = _tool.Execute(operation, pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);
        Assert.StartsWith("Hyperlink deleted from slide", result);
    }

    [Theory]
    [InlineData("GET")]
    [InlineData("Get")]
    [InlineData("get")]
    public void Operation_ShouldBeCaseInsensitive_Get(string operation)
    {
        var pptPath = CreatePresentationWithHyperlink($"test_case_get_{operation}.pptx");
        var result = _tool.Execute(operation, pptPath);
        Assert.Contains("Hyperlink", result, StringComparison.OrdinalIgnoreCase);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
        Assert.Contains("Unknown operation", ex.Message);
    }

    [Fact]
    public void Add_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_no_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", pptPath, url: "https://example.com"));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void Add_WithoutUrlOrSlideTargetIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_no_target.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", pptPath, slideIndex: 0, text: "Link"));
        Assert.Contains("url or slideTargetIndex", ex.Message);
    }

    [Fact]
    public void Add_WithLinkTextNotFound_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_add_linktext_notfound.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", pptPath, slideIndex: 0,
            url: "https://example.com", text: "Some text", linkText: "notfound"));
        Assert.Contains("linkText", ex.Message);
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public void Edit_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithHyperlink("test_edit_no_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pptPath, shapeIndex: 0, url: "https://new.com"));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void Edit_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithHyperlink("test_edit_no_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("edit", pptPath, slideIndex: 0, url: "https://new.com"));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void Delete_WithoutSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithHyperlink("test_delete_no_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", pptPath, shapeIndex: 0));
        Assert.Contains("slideIndex is required", ex.Message);
    }

    [Fact]
    public void Delete_WithoutShapeIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreatePresentationWithHyperlink("test_delete_no_shape.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("delete", pptPath, slideIndex: 0));
        Assert.Contains("shapeIndex is required", ex.Message);
    }

    [Fact]
    public void Get_WithInvalidSlideIndex_ShouldThrowArgumentException()
    {
        var pptPath = CreateTestPresentation("test_get_invalid_slide.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("get", pptPath, slideIndex: 99));
        Assert.Contains("slideIndex", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void Get_WithSessionId_ShouldReturnHyperlinks()
    {
        var pptPath = CreatePresentationWithHyperlink("test_session_get_hyperlinks.pptx", "https://session-test.com");
        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.Contains("Hyperlink", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Add_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add_hyperlink.pptx");
        var shapeIndex = FindShapeIndex(pptPath);
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, shapeIndex: shapeIndex,
            url: "https://session-example.com", text: "Session Link");
        Assert.StartsWith("Hyperlink added to slide", result);
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
        Assert.StartsWith("Hyperlink updated on slide", result);
        Assert.Contains("session-new.com", ppt.Slides[0].Shapes[0].HyperlinkClick?.ExternalUrl ?? "");
    }

    [Fact]
    public void Delete_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreatePresentationWithHyperlink("test_session_delete_hyperlink.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0, shapeIndex: 0);
        Assert.StartsWith("Hyperlink deleted from slide", result);
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
        var pptPath1 = CreateTestPresentation("test_path_hyperlink.pptx");
        var pptPath2 = CreatePresentationWithHyperlink("test_session_hyperlink.pptx", "https://session-url.com");
        var sessionId = OpenSession(pptPath2);
        var result = _tool.Execute("get", pptPath1, sessionId);
        Assert.Contains("session-url.com", result);
    }

    #endregion
}