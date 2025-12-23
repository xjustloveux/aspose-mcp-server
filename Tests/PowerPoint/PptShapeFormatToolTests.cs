using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptShapeFormatToolTests : TestBase
{
    private readonly PptShapeFormatTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        // Use the default first slide instead of AddEmptySlide to ensure shapes are properly saved
        var slide = presentation.Slides[0];
        slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task SetFill_ShouldSetShapeFill()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_fill.pptx");

        // Find the correct shapeIndex for the added AutoShape (excluding placeholders)
        var correctShapeIndex = -1;
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var nonPlaceholderShapes = pptSlide.Shapes.Where(s => s.Placeholder == null).ToList();
            Assert.True(nonPlaceholderShapes.Count > 0, "Should find at least one non-placeholder shape");
            // The added shape should be the one with original coordinates (100, 100)
            for (var i = 0; i < nonPlaceholderShapes.Count; i++)
            {
                var s = nonPlaceholderShapes[i];
                if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                {
                    correctShapeIndex = pptSlide.Shapes.IndexOf(s);
                    break;
                }
            }

            if (correctShapeIndex < 0)
                correctShapeIndex =
                    pptSlide.Shapes.IndexOf(nonPlaceholderShapes[0]); // Fallback to first non-placeholder shape
        }

        Assert.True(correctShapeIndex >= 0, "Should find at least one non-placeholder shape");

        var outputPath = CreateTestFilePath("test_set_fill_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_fill",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = correctShapeIndex,
            ["fillType"] = "Solid",
            ["color"] = "#FF0000"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }

    [Fact]
    public async Task SetLine_ShouldSetShapeLine()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_line.pptx");

        // Find the correct shapeIndex for the added AutoShape (excluding placeholders)
        var correctShapeIndex = -1;
        using (var ppt = new Presentation(pptPath))
        {
            var pptSlide = ppt.Slides[0];
            var nonPlaceholderShapes = pptSlide.Shapes.Where(s => s.Placeholder == null).ToList();
            Assert.True(nonPlaceholderShapes.Count > 0, "Should find at least one non-placeholder shape");
            // The added shape should be the one with original coordinates (100, 100)
            for (var i = 0; i < nonPlaceholderShapes.Count; i++)
            {
                var s = nonPlaceholderShapes[i];
                if (Math.Abs(s.X - 100) < 1 && Math.Abs(s.Y - 100) < 1)
                {
                    correctShapeIndex = pptSlide.Shapes.IndexOf(s);
                    break;
                }
            }

            if (correctShapeIndex < 0)
                correctShapeIndex =
                    pptSlide.Shapes.IndexOf(nonPlaceholderShapes[0]); // Fallback to first non-placeholder shape
        }

        Assert.True(correctShapeIndex >= 0, "Should find at least one non-placeholder shape");

        var outputPath = CreateTestFilePath("test_set_line_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_line",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = correctShapeIndex,
            ["color"] = "#0000FF",
            ["width"] = 2.0
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }
}