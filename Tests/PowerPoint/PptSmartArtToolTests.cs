using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptSmartArtToolTests : TestBase
{
    private readonly PptSmartArtTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task AddSmartArt_ShouldAddSmartArt()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_add_smartart.pptx");
        var outputPath = CreateTestFilePath("test_add_smartart_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["layout"] = "BasicProcess",
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 400,
            ["height"] = 300
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        var smartArts = slide.Shapes.OfType<ISmartArt>().ToList();
        Assert.True(smartArts.Count > 0, "Slide should contain at least one SmartArt");
    }

    [Fact]
    public async Task GetSmartArt_ShouldReturnSmartArtInfo()
    {
        // Arrange - PptSmartArtTool doesn't have a "get" operation, test manage_nodes instead
        var pptPath = CreateTestPresentation("test_manage_smartart_nodes.pptx");
        int smartArtShapeIndex;
        using (var presentation = new Presentation(pptPath))
        {
            var slide = presentation.Slides[0];
            var smartArt = slide.Shapes.AddSmartArt(100, 100, 400, 300, SmartArtLayoutType.BasicProcess);
            presentation.Save(pptPath, SaveFormat.Pptx);

            // Find the index of the SmartArt shape (it might not be at index 0 if there are placeholders)
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
        var arguments = new JsonObject
        {
            ["operation"] = "manage_nodes",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["shapeIndex"] = smartArtShapeIndex,
            ["action"] = "add",
            ["targetPath"] = new JsonArray { 0 },
            ["text"] = "New Node"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("SmartArt", result, StringComparison.OrdinalIgnoreCase);
    }
}