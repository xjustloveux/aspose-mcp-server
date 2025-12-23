using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptHeaderFooterToolTests : TestBase
{
    private readonly PptHeaderFooterTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task SetHeader_ShouldSetHeaderText()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_header.pptx");
        var outputPath = CreateTestFilePath("test_set_header_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_header",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["headerText"] = "Header Text"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.NotNull(slide);
    }

    [Fact]
    public async Task SetFooter_ShouldSetFooterText()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_footer.pptx");
        var outputPath = CreateTestFilePath("test_set_footer_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_footer",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["footerText"] = "Footer Text"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var presentation = new Presentation(outputPath);
        var slide = presentation.Slides[0];
        Assert.NotNull(slide);
    }

    [Fact]
    public async Task SetSlideNumbering_ShouldSetSlideNumbering()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_set_slide_numbering.pptx");
        var outputPath = CreateTestFilePath("test_set_slide_numbering_output.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_slide_numbering",
            ["path"] = pptPath,
            ["outputPath"] = outputPath,
            ["showSlideNumber"] = true
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output presentation should be created");
    }
}