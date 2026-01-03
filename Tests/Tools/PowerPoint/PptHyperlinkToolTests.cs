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
        // Use the default first slide instead of AddEmptySlide to ensure shapes are properly saved
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    #region General Tests

    [Fact]
    public void AddHyperlink_ShouldAddHyperlink()
    {
        var pptPath = CreateTestPresentation("test_add_hyperlink.pptx");

        // Find the correct shapeIndex for the added AutoShape (excluding placeholders)
        var correctShapeIndex = -1;
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var nonPlaceholderShapes = pptSlide.Shapes.Where(s => s.Placeholder == null).ToList();
            // If no non-placeholder shapes found, use all shapes
            if (nonPlaceholderShapes.Count == 0) nonPlaceholderShapes = pptSlide.Shapes.ToList();
            Assert.True(nonPlaceholderShapes.Count > 0,
                $"Should find at least one shape. Total shapes: {pptSlide.Shapes.Count}, Non-placeholder: {pptSlide.Shapes.Count(s => s.Placeholder == null)}");
            // The added shape should be the one with original coordinates (100, 100)
            foreach (var s in nonPlaceholderShapes)
                if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                {
                    correctShapeIndex = pptSlide.Shapes.IndexOf(s);
                    break;
                }

            if (correctShapeIndex < 0)
                correctShapeIndex =
                    pptSlide.Shapes.IndexOf(nonPlaceholderShapes[0]); // Fallback to first non-placeholder shape
        }

        Assert.True(correctShapeIndex >= 0, $"Should find at least one shape. Found shape index: {correctShapeIndex}");

        var outputPath = CreateTestFilePath("test_add_hyperlink_output.pptx");
        _tool.Execute("add", pptPath, slideIndex: 0, shapeIndex: correctShapeIndex, url: "https://example.com",
            text: "Click here", outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.True(correctShapeIndex < slide.Shapes.Count, $"Shape index {correctShapeIndex} should be valid");
        var shape = slide.Shapes[correctShapeIndex];
        Assert.NotNull(shape.HyperlinkClick);
        Assert.Contains("example.com", shape.HyperlinkClick.ExternalUrl ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GetHyperlinks_ShouldReturnAllHyperlinks()
    {
        var pptPath = CreateTestPresentation("test_get_hyperlinks.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            // Find non-placeholder shape or use first shape
            var shape = slide.Shapes.OfType<IAutoShape>().FirstOrDefault(s => s.Placeholder == null)
                        ?? slide.Shapes.OfType<IAutoShape>().FirstOrDefault()
                        ?? slide.Shapes[0];
            shape.HyperlinkClick = new Hyperlink("https://test.com");
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", pptPath);
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Hyperlink", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EditHyperlink_ShouldModifyHyperlink()
    {
        var pptPath = CreateTestPresentation("test_edit_hyperlink.pptx");

        // Find the correct shapeIndex for the added AutoShape (excluding placeholders)
        var correctShapeIndex = -1;
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var nonPlaceholderShapes = pptSlide.Shapes.Where(s => s.Placeholder == null).ToList();
            // If no non-placeholder shapes found, use all shapes
            if (nonPlaceholderShapes.Count == 0) nonPlaceholderShapes = pptSlide.Shapes.ToList();
            Assert.True(nonPlaceholderShapes.Count > 0,
                $"Should find at least one shape. Total shapes: {pptSlide.Shapes.Count}, Non-placeholder: {pptSlide.Shapes.Count(s => s.Placeholder == null)}");
            // The added shape should be the one with original coordinates (100, 100)
            foreach (var s in nonPlaceholderShapes)
                if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                {
                    correctShapeIndex = pptSlide.Shapes.IndexOf(s);
                    break;
                }

            if (correctShapeIndex < 0)
                correctShapeIndex =
                    pptSlide.Shapes.IndexOf(nonPlaceholderShapes[0]); // Fallback to first non-placeholder shape

            // Set initial hyperlink
            var pptShape = pptSlide.Shapes[correctShapeIndex];
            pptShape.HyperlinkClick = new Hyperlink("https://old.com");
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        Assert.True(correctShapeIndex >= 0, $"Should find at least one shape. Found shape index: {correctShapeIndex}");

        var outputPath = CreateTestFilePath("test_edit_hyperlink_output.pptx");
        _tool.Execute("edit", pptPath, slideIndex: 0, shapeIndex: correctShapeIndex, url: "https://new.com",
            outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.True(correctShapeIndex < slide.Shapes.Count, $"Shape index {correctShapeIndex} should be valid");
        var shape = slide.Shapes[correctShapeIndex];
        Assert.Contains("new.com", shape.HyperlinkClick?.ExternalUrl ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DeleteHyperlink_ShouldDeleteHyperlink()
    {
        var pptPath = CreateTestPresentation("test_delete_hyperlink.pptx");

        // Find the correct shapeIndex for the added AutoShape (excluding placeholders)
        var correctShapeIndex = -1;
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var nonPlaceholderShapes = pptSlide.Shapes.Where(s => s.Placeholder == null).ToList();
            // If no non-placeholder shapes found, use all shapes
            if (nonPlaceholderShapes.Count == 0) nonPlaceholderShapes = pptSlide.Shapes.ToList();
            Assert.True(nonPlaceholderShapes.Count > 0,
                $"Should find at least one shape. Total shapes: {pptSlide.Shapes.Count}, Non-placeholder: {pptSlide.Shapes.Count(s => s.Placeholder == null)}");
            // The added shape should be the one with original coordinates (100, 100)
            foreach (var s in nonPlaceholderShapes)
                if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                {
                    correctShapeIndex = pptSlide.Shapes.IndexOf(s);
                    break;
                }

            if (correctShapeIndex < 0)
                correctShapeIndex =
                    pptSlide.Shapes.IndexOf(nonPlaceholderShapes[0]); // Fallback to first non-placeholder shape

            // Set initial hyperlink
            var pptShape = pptSlide.Shapes[correctShapeIndex];
            pptShape.HyperlinkClick = new Hyperlink("https://delete.com");
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        Assert.True(correctShapeIndex >= 0, $"Should find at least one shape. Found shape index: {correctShapeIndex}");

        var outputPath = CreateTestFilePath("test_delete_hyperlink_output.pptx");
        _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: correctShapeIndex, outputPath: outputPath);
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.True(correctShapeIndex < slide.Shapes.Count, $"Shape index {correctShapeIndex} should be valid");
        var shape = slide.Shapes[correctShapeIndex];
        Assert.Null(shape.HyperlinkClick);
    }

    [Fact]
    public void AddHyperlink_WithSlideTargetIndex_ShouldAddInternalLink()
    {
        // Arrange - Create presentation with multiple slides
        var pptPath = CreateTestFilePath("test_add_internal_link.pptx");
        using (var ppt = new Presentation())
        {
            ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
            ppt.Slides.AddEmptySlide(ppt.LayoutSlides[0]);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_add_internal_link_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, slideTargetIndex: 1, text: "Go to slide 2",
            outputPath: outputPath);
        Assert.Contains("Slide 1", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void AddHyperlink_WithNewShape_ShouldCreateShapeWithLink()
    {
        var pptPath = CreateTestPresentation("test_add_new_shape.pptx");
        var outputPath = CreateTestFilePath("test_add_new_shape_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, url: "https://example.com", text: "New Link", x: 100,
            y: 200, width: 150, height: 40, outputPath: outputPath);
        Assert.Contains("Hyperlink added", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void GetHyperlinks_WithSlideIndex_ShouldReturnSlideHyperlinks()
    {
        var pptPath = CreateTestPresentation("test_get_slide_hyperlinks.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            var shape = slide.Shapes.OfType<IAutoShape>().FirstOrDefault(s => s.Placeholder == null)
                        ?? slide.Shapes.OfType<IAutoShape>().FirstOrDefault();
            if (shape != null)
                shape.HyperlinkClick = new Hyperlink("https://test.com");
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        Assert.Contains("slideIndex", result);
        Assert.Contains("hyperlinks", result);
    }

    [SkippableFact]
    public void AddHyperlink_WithLinkText_ShouldAddPortionLevelHyperlink()
    {
        // Skip in evaluation mode - evaluation watermark interferes with text content
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation watermark interferes with text content");
        var pptPath = CreateTestPresentation("test_add_portion_hyperlink.pptx");
        var outputPath = CreateTestFilePath("test_add_portion_hyperlink_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, url: "https://example.com",
            text: "Please click here for more info", linkText: "here", outputPath: outputPath);
        Assert.Contains("Hyperlink added", result);
        Assert.Contains("on text: 'here'", result);
        Assert.True(File.Exists(outputPath));

        // Verify the hyperlink is on the correct portion
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var autoShape = slide.Shapes.OfType<IAutoShape>().Last();
        Assert.NotNull(autoShape.TextFrame);

        var allText = string.Join("", autoShape.TextFrame.Paragraphs
            .SelectMany(p => p.Portions)
            .Select(p => p.Text));
        Assert.Equal("Please click here for more info", allText);

        // Find the "here" portion and verify it has the hyperlink
        var herePortion = autoShape.TextFrame.Paragraphs
            .SelectMany(p => p.Portions)
            .FirstOrDefault(p => p.Text == "here");
        Assert.NotNull(herePortion);
        Assert.NotNull(herePortion.PortionFormat.HyperlinkClick);
        Assert.Contains("example.com", herePortion.PortionFormat.HyperlinkClick.ExternalUrl ?? "");
    }

    [Fact]
    public void GetHyperlinks_ShouldReturnPortionLevelHyperlinks()
    {
        // Arrange - Create presentation with portion-level hyperlink
        var pptPath = CreateTestFilePath("test_get_portion_hyperlinks.pptx");
        using (var ppt = new Presentation())
        {
            var slide = ppt.Slides[0];
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
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var result = _tool.Execute("get", pptPath, slideIndex: 0);
        Assert.Contains("\"level\": \"text\"", result);
        Assert.Contains("\"text\": \"here\"", result);
        Assert.Contains("portion-link.com", result);
    }

    [Fact]
    public void DeleteHyperlink_ShouldDeletePortionLevelHyperlinks()
    {
        // Arrange - Create presentation with portion-level hyperlink
        var pptPath = CreateTestFilePath("test_delete_portion_hyperlink.pptx");
        using (var ppt = new Presentation())
        {
            var pptSlide = ppt.Slides[0];
            var shape = pptSlide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 50);
            shape.TextFrame.Paragraphs.Clear();
            var paragraph = new Paragraph();
            var linkPortion = new Portion("Click here")
            {
                PortionFormat = { HyperlinkClick = new Hyperlink("https://delete-me.com") }
            };
            paragraph.Portions.Add(linkPortion);
            shape.TextFrame.Paragraphs.Add(paragraph);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_portion_hyperlink_output.pptx");
        _tool.Execute("delete", pptPath, slideIndex: 0, shapeIndex: 0, outputPath: outputPath);

        // Assert - Verify portion-level hyperlink is deleted
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var autoShape = slide.Shapes.OfType<IAutoShape>().First();
        foreach (var para in autoShape.TextFrame.Paragraphs)
        foreach (var portion in para.Portions)
            Assert.Null(portion.PortionFormat.HyperlinkClick);
    }

    [SkippableFact]
    public void AddHyperlink_WithLinkTextAtStart_ShouldAddPortionLevelHyperlink()
    {
        // Skip in evaluation mode - evaluation watermark interferes with text content
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation watermark interferes with text content");

        // Arrange - linkText at the beginning of text
        var pptPath = CreateTestPresentation("test_add_linktext_start.pptx");
        var outputPath = CreateTestFilePath("test_add_linktext_start_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, url: "https://example.com",
            text: "Click here for more", linkText: "Click", outputPath: outputPath);
        Assert.Contains("on text: 'Click'", result);

        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var autoShape = slide.Shapes.OfType<IAutoShape>().Last();
        var clickPortion = autoShape.TextFrame.Paragraphs
            .SelectMany(p => p.Portions)
            .FirstOrDefault(p => p.Text == "Click");
        Assert.NotNull(clickPortion);
        Assert.NotNull(clickPortion.PortionFormat.HyperlinkClick);
    }

    [SkippableFact]
    public void AddHyperlink_WithLinkTextAtEnd_ShouldAddPortionLevelHyperlink()
    {
        // Skip in evaluation mode - evaluation watermark interferes with text content
        SkipInEvaluationMode(AsposeLibraryType.Slides, "Evaluation watermark interferes with text content");

        // Arrange - linkText at the end of text
        var pptPath = CreateTestPresentation("test_add_linktext_end.pptx");
        var outputPath = CreateTestFilePath("test_add_linktext_end_output.pptx");
        var result = _tool.Execute("add", pptPath, slideIndex: 0, url: "https://example.com", text: "More info here",
            linkText: "here", outputPath: outputPath);
        Assert.Contains("on text: 'here'", result);

        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var autoShape = slide.Shapes.OfType<IAutoShape>().Last();
        var herePortion = autoShape.TextFrame.Paragraphs
            .SelectMany(p => p.Portions)
            .FirstOrDefault(p => p.Text == "here");
        Assert.NotNull(herePortion);
        Assert.NotNull(herePortion.PortionFormat.HyperlinkClick);
    }

    #endregion

    #region Exception Tests

    [Fact]
    public void ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("unknown", pptPath));
    }

    [Fact]
    public void AddHyperlink_WithoutUrlOrSlideTargetIndex_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_add_no_target.pptx");
        Assert.Throws<ArgumentException>(() => _tool.Execute("add", pptPath, slideIndex: 0, text: "Link"));
    }

    [Fact]
    public void AddHyperlink_WithLinkTextNotFound_ShouldThrow()
    {
        var pptPath = CreateTestPresentation("test_add_linktext_notfound.pptx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("add", pptPath, slideIndex: 0,
            url: "https://example.com", text: "Some text without the link word", linkText: "notfound"));
        Assert.Contains("linkText", ex.Message);
        Assert.Contains("not found", ex.Message);
    }

    #endregion

    #region Session ID Tests

    [Fact]
    public void GetHyperlinks_WithSessionId_ShouldReturnHyperlinks()
    {
        var pptPath = CreateTestPresentation("test_session_get_hyperlinks.pptx");
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            var shape = slide.Shapes.OfType<IAutoShape>().FirstOrDefault(s => s.Placeholder == null)
                        ?? slide.Shapes.OfType<IAutoShape>().FirstOrDefault()
                        ?? slide.Shapes[0];
            shape.HyperlinkClick = new Hyperlink("https://session-test.com");
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var result = _tool.Execute("get", sessionId: sessionId);
        Assert.NotNull(result);
        Assert.Contains("Hyperlink", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AddHyperlink_WithSessionId_ShouldAddInMemory()
    {
        var pptPath = CreateTestPresentation("test_session_add_hyperlink.pptx");
        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);

        // Find correct shape index
        var slide = ppt.Slides[0];
        var nonPlaceholderShapes = slide.Shapes.Where(s => s.Placeholder == null).ToList();
        if (nonPlaceholderShapes.Count == 0) nonPlaceholderShapes = slide.Shapes.ToList();
        var correctShapeIndex = slide.Shapes.IndexOf(nonPlaceholderShapes[0]);
        var result = _tool.Execute("add", sessionId: sessionId, slideIndex: 0, shapeIndex: correctShapeIndex,
            url: "https://session-example.com", text: "Session Link");
        Assert.Contains("Hyperlink added", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var shape = ppt.Slides[0].Shapes[correctShapeIndex];
        Assert.NotNull(shape.HyperlinkClick);
        Assert.Contains("session-example.com", shape.HyperlinkClick.ExternalUrl ?? "");
    }

    [Fact]
    public void DeleteHyperlink_WithSessionId_ShouldDeleteInMemory()
    {
        var pptPath = CreateTestFilePath("test_session_delete_hyperlink.pptx");
        using (var presentation = new Presentation())
        {
            var slideToSetup = presentation.Slides[0];
            var shapeToSetup = slideToSetup.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 50);
            shapeToSetup.HyperlinkClick = new Hyperlink("https://to-delete.com");
            presentation.Save(pptPath, SaveFormat.Pptx);
        }

        var sessionId = OpenSession(pptPath);
        var ppt = SessionManager.GetDocument<Presentation>(sessionId);
        var result = _tool.Execute("delete", sessionId: sessionId, slideIndex: 0, shapeIndex: 0);
        Assert.Contains("Hyperlink deleted", result);
        Assert.Contains("session", result);

        // Verify in-memory changes
        var shape = ppt.Slides[0].Shapes[0];
        Assert.Null(shape.HyperlinkClick);
    }

    #endregion
}