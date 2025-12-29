using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptHyperlinkToolTests : TestBase
{
    private readonly PptHyperlinkTool _tool = new();

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

    [Fact]
    public async Task AddHyperlink_ShouldAddHyperlink()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = correctShapeIndex,
            ["url"] = "https://example.com",
            ["text"] = "Click here"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.True(correctShapeIndex < slide.Shapes.Count, $"Shape index {correctShapeIndex} should be valid");
        var shape = slide.Shapes[correctShapeIndex];
        Assert.NotNull(shape.HyperlinkClick);
        Assert.Contains("example.com", shape.HyperlinkClick.ExternalUrl ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetHyperlinks_ShouldReturnAllHyperlinks()
    {
        // Arrange
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

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Hyperlink", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task EditHyperlink_ShouldModifyHyperlink()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "edit",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = correctShapeIndex,
            ["url"] = "https://new.com"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.True(correctShapeIndex < slide.Shapes.Count, $"Shape index {correctShapeIndex} should be valid");
        var shape = slide.Shapes[correctShapeIndex];
        Assert.Contains("new.com", shape.HyperlinkClick?.ExternalUrl ?? "", StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task DeleteHyperlink_ShouldDeleteHyperlink()
    {
        // Arrange
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
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = correctShapeIndex
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.True(correctShapeIndex < slide.Shapes.Count, $"Shape index {correctShapeIndex} should be valid");
        var shape = slide.Shapes[correctShapeIndex];
        Assert.Null(shape.HyperlinkClick);
    }

    [Fact]
    public async Task AddHyperlink_WithSlideTargetIndex_ShouldAddInternalLink()
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
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["slideTargetIndex"] = 1,
            ["text"] = "Go to slide 2"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Slide 1", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task AddHyperlink_WithoutUrlOrSlideTargetIndex_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_no_target.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["text"] = "Link"
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task AddHyperlink_WithNewShape_ShouldCreateShapeWithLink()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_new_shape.pptx");
        var outputPath = CreateTestFilePath("test_add_new_shape_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["url"] = "https://example.com",
            ["text"] = "New Link",
            ["x"] = 100,
            ["y"] = 200,
            ["width"] = 150,
            ["height"] = 40
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Hyperlink added", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task GetHyperlinks_WithSlideIndex_ShouldReturnSlideHyperlinks()
    {
        // Arrange
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

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("slideIndex", result);
        Assert.Contains("hyperlinks", result);
    }

    [Fact]
    public async Task AddHyperlink_WithLinkText_ShouldAddPortionLevelHyperlink()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_portion_hyperlink.pptx");
        var outputPath = CreateTestFilePath("test_add_portion_hyperlink_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["url"] = "https://example.com",
            ["text"] = "Please click here for more info",
            ["linkText"] = "here"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
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
    public async Task AddHyperlink_WithLinkTextNotFound_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_linktext_notfound.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["slideIndex"] = 0,
            ["url"] = "https://example.com",
            ["text"] = "Some text without the link word",
            ["linkText"] = "notfound"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("linkText", ex.Message);
        Assert.Contains("not found", ex.Message);
    }

    [Fact]
    public async Task GetHyperlinks_ShouldReturnPortionLevelHyperlinks()
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
            var linkPortion = new Portion("here");
            linkPortion.PortionFormat.HyperlinkClick = new Hyperlink("https://portion-link.com");
            paragraph.Portions.Add(linkPortion);
            paragraph.Portions.Add(new Portion(" for more"));
            shape.TextFrame.Paragraphs.Add(paragraph);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"level\": \"text\"", result);
        Assert.Contains("\"text\": \"here\"", result);
        Assert.Contains("portion-link.com", result);
    }

    [Fact]
    public async Task DeleteHyperlink_ShouldDeletePortionLevelHyperlinks()
    {
        // Arrange - Create presentation with portion-level hyperlink
        var pptPath = CreateTestFilePath("test_delete_portion_hyperlink.pptx");
        using (var ppt = new Presentation())
        {
            var pptSlide = ppt.Slides[0];
            var shape = pptSlide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 300, 50);
            shape.TextFrame.Paragraphs.Clear();
            var paragraph = new Paragraph();
            var linkPortion = new Portion("Click here");
            linkPortion.PortionFormat.HyperlinkClick = new Hyperlink("https://delete-me.com");
            paragraph.Portions.Add(linkPortion);
            shape.TextFrame.Paragraphs.Add(paragraph);
            ppt.Save(pptPath, SaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_delete_portion_hyperlink_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = 0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert - Verify portion-level hyperlink is deleted
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var autoShape = slide.Shapes.OfType<IAutoShape>().First();
        foreach (var para in autoShape.TextFrame.Paragraphs)
        foreach (var portion in para.Portions)
            Assert.Null(portion.PortionFormat.HyperlinkClick);
    }

    [Fact]
    public async Task AddHyperlink_WithLinkTextAtStart_ShouldAddPortionLevelHyperlink()
    {
        // Arrange - linkText at the beginning of text
        var pptPath = CreateTestPresentation("test_add_linktext_start.pptx");
        var outputPath = CreateTestFilePath("test_add_linktext_start_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["url"] = "https://example.com",
            ["text"] = "Click here for more",
            ["linkText"] = "Click"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
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

    [Fact]
    public async Task AddHyperlink_WithLinkTextAtEnd_ShouldAddPortionLevelHyperlink()
    {
        // Arrange - linkText at the end of text
        var pptPath = CreateTestPresentation("test_add_linktext_end.pptx");
        var outputPath = CreateTestFilePath("test_add_linktext_end_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["url"] = "https://example.com",
            ["text"] = "More info here",
            ["linkText"] = "here"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
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
}