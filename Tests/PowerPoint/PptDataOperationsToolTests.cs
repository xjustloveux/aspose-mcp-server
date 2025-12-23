using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptDataOperationsToolTests : TestBase
{
    private readonly PptDataOperationsTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task GetStatistics_ShouldReturnStatistics()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_statistics.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_statistics",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Statistics", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetContent_ShouldReturnContent()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_content.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_content",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        // Content may be empty for blank slides, just verify result is returned
        Assert.True(result.Length > 0, "Result should not be empty");
    }

    [Fact]
    public async Task GetSlideDetails_ShouldReturnSlideDetails()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_slide_details.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_slide_details",
            ["path"] = pptPath,
            ["slideIndex"] = 0
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Slide", result, StringComparison.OrdinalIgnoreCase);
    }
}