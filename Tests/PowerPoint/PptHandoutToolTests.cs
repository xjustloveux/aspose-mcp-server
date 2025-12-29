using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.PowerPoint;

namespace AsposeMcpServer.Tests.PowerPoint;

public class PptHandoutToolTests : TestBase
{
    private readonly PptHandoutTool _tool = new();

    private string CreateTestPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        presentation.Save(filePath, SaveFormat.Pptx);
        return filePath;
    }

    [Fact]
    public async Task SetHeaderFooter_WithoutHandoutMaster_ShouldThrowWithHelpfulMessage()
    {
        // Arrange - New presentations don't have handout master by default
        var pptPath = CreateTestPresentation("test_handout_no_master.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_header_footer",
            ["path"] = pptPath,
            ["headerText"] = "Handout Header"
        };

        // Act & Assert
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("PowerPoint", ex.Message);
    }

    [Fact]
    public async Task SetHeaderFooter_ErrorMessage_ShouldContainInstructions()
    {
        // Arrange - Verify the error message contains helpful instructions
        var pptPath = CreateTestPresentation("test_handout_instructions.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_header_footer",
            ["path"] = pptPath,
            ["headerText"] = "Header"
        };

        // Act
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => _tool.ExecuteAsync(arguments));

        // Assert - Error message should contain instructions for creating handout master
        Assert.Contains("View", ex.Message);
        Assert.Contains("Handout Master", ex.Message);
    }

    [Fact]
    public async Task SetHeaderFooter_WithAllParameters_ShouldThrowWithoutMaster()
    {
        // Arrange - Test with all parameters to ensure they are parsed correctly before the error
        var pptPath = CreateTestPresentation("test_handout_all_params.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "set_header_footer",
            ["path"] = pptPath,
            ["headerText"] = "Header",
            ["footerText"] = "Footer",
            ["dateText"] = "2024-12-28",
            ["showPageNumber"] = true
        };

        // Act & Assert - Should throw because no handout master, but parameters should be valid
        var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("handout master", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ExecuteAsync_UnknownOperation_ShouldThrow()
    {
        // Arrange
        var pptPath = CreateTestPresentation("test_unknown_op.pptx");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pptPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}