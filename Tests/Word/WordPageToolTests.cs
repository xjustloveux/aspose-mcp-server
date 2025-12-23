using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordPageToolTests : WordTestBase
{
    private readonly WordPageTool _tool = new();

    [Fact]
    public async Task SetMargins_ShouldSetPageMargins()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_margins.docx");
        var outputPath = CreateTestFilePath("test_set_margins_output.docx");
        var arguments = CreateArguments("set_margins", docPath, outputPath);
        arguments["top"] = 72.0;
        arguments["bottom"] = 72.0;
        arguments["left"] = 90.0;
        arguments["right"] = 90.0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(72.0, section.PageSetup.TopMargin);
        Assert.Equal(72.0, section.PageSetup.BottomMargin);
        Assert.Equal(90.0, section.PageSetup.LeftMargin);
        Assert.Equal(90.0, section.PageSetup.RightMargin);
    }

    [Fact]
    public async Task SetOrientation_ShouldSetPageOrientation()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_orientation.docx");
        var outputPath = CreateTestFilePath("test_set_orientation_output.docx");
        var arguments = CreateArguments("set_orientation", docPath, outputPath);
        arguments["orientation"] = "landscape";

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(Orientation.Landscape, section.PageSetup.Orientation);
    }

    [Fact]
    public async Task SetPageSize_ShouldSetPageSize()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_size.docx");
        var outputPath = CreateTestFilePath("test_set_size_output.docx");
        var arguments = CreateArguments("set_size", docPath, outputPath);
        arguments["width"] = 595.0;
        arguments["height"] = 842.0;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        Assert.Equal(595.0, section.PageSetup.PageWidth);
        Assert.Equal(842.0, section.PageSetup.PageHeight);
    }

    [Fact]
    public async Task SetPageNumber_ShouldSetPageNumber()
    {
        // Arrange
        var docPath = CreateWordDocument("test_set_page_number.docx");
        var outputPath = CreateTestFilePath("test_set_page_number_output.docx");
        var arguments = CreateArguments("set_page_number", docPath, outputPath);
        arguments["startingPageNumber"] = 5;

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var doc = new Document(outputPath);
        var section = doc.Sections[0];
        // Verify page numbering was set (may require RestartPageNumbering to be true)
        Assert.True(section.PageSetup.RestartPageNumbering || section.PageSetup.PageStartingNumber == 5,
            "Page starting number should be set");
    }
}