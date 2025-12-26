using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptPropertiesToolTests : TestBase
{
    private readonly PptPropertiesTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task GetProperties_ShouldReturnProperties()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_properties.pptx");
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
        Assert.Contains("title", result);
    }

    [Fact]
    public async Task SetProperties_ShouldSetProperties()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_properties.pptx");
        var outputPath = CreateTestFilePath("test_set_properties_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["title"] = "Test Presentation",
            ["author"] = "Test Author",
            ["subject"] = "Test Subject"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        Assert.Equal("Test Presentation", presentation.DocumentProperties.Title);
        Assert.Equal("Test Author", presentation.DocumentProperties.Author);
        Assert.Equal("Test Subject", presentation.DocumentProperties.Subject);
    }
}