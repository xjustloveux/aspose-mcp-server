using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptMediaToolTests : TestBase
{
    private readonly PptMediaTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task GetMedia_ShouldReturnMediaInfo()
    {
        // Arrange - PptMediaTool doesn't have a "get" operation, test add_audio instead
        var pptPath = CreateTestPresentation("test_add_audio.pptx");
        var audioPath = CreateTestFilePath("test_audio.mp3");
        File.WriteAllText(audioPath, "fake audio content");

        var outputPath = CreateTestFilePath("test_add_audio_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "add_audio",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["audioPath"] = audioPath,
            ["x"] = 100,
            ["y"] = 100
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Audio", result, StringComparison.OrdinalIgnoreCase);
    }
}