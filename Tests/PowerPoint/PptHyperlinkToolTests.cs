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
            ["displayText"] = "Click here"
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
}