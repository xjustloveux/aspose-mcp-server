using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptBackgroundToolTests : TestBase
{
    private readonly PptBackgroundTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task SetBackgroundColor_ShouldSetBackgroundColor()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_background_color.pptx");
        var outputPath = CreateTestFilePath("test_set_background_color_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["color"] = "#FF0000"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.NotNull(slide.Background);
    }

    [Fact]
    public async Task GetBackground_ShouldReturnBackgroundInfo()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_background.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Background", result, StringComparison.OrdinalIgnoreCase);
    }
}