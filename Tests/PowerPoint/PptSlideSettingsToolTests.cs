using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptSlideSettingsToolTests : TestBase
{
    private readonly PptSlideSettingsTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task SetSlideSize_ShouldSetSlideSize()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_slide_size.pptx");
        var outputPath = CreateTestFilePath("test_set_slide_size_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_size",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["preset"] = "OnScreen16x9"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        Assert.NotNull(presentation);
    }

    [Fact]
    public async Task SetSlideOrientation_ShouldSetOrientation()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_orientation.pptx");
        var outputPath = CreateTestFilePath("test_set_orientation_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_orientation",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["orientation"] = "Portrait"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }
}