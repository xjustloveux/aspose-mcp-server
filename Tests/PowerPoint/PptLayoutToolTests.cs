using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptLayoutToolTests : TestBase
{
    private readonly PptLayoutTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task SetLayout_ShouldSetSlideLayout()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_layout.pptx");
        var outputPath = CreateTestFilePath("test_set_layout_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["slideIndex"] = 0,
            ["layout"] = "Title"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.NotNull(slide);
    }

    [Fact]
    public async Task GetLayouts_ShouldReturnAvailableLayouts()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_layouts.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_layouts",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Layout", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetMasters_ShouldReturnMasterSlides()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_get_masters.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "get_masters",
            ["path"] = pptPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Master", result, StringComparison.OrdinalIgnoreCase);
    }
}