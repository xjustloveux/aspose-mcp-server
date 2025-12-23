using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Word;

namespace AsposeMcpServer.Tests.Word;

public class WordContentToolTests : WordTestBase
{
    private readonly WordContentTool _tool = new();

    [Fact]
    public async Task GetContent_ShouldReturnContent()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_get_content.docx", "Test content for extraction");
        var arguments = CreateArguments("get_content", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("content", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetContentDetailed_ShouldReturnDetailedContent()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_get_content_detailed.docx", "Detailed content");
        var arguments = CreateArguments("get_content_detailed", docPath);
        arguments["includeFooters"] = true;
        arguments["includeHeaders"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public async Task GetStatistics_ShouldReturnStatistics()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_get_statistics.docx", "Test document for statistics");
        var arguments = CreateArguments("get_statistics", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Statistics", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetDocumentInfo_ShouldReturnDocumentInfo()
    {
        // Arrange
        var docPath = CreateWordDocument("test_get_document_info.docx");
        var arguments = CreateArguments("get_document_info", docPath);

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Document", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetStatistics_WithIncludeFootnotes_ShouldRespectParameter()
    {
        // Arrange
        var docPath = CreateWordDocumentWithContent("test_statistics_footnotes.docx", "Test content");
        var arguments = CreateArguments("get_statistics", docPath);
        arguments["includeFootnotes"] = false;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("excluded", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetDocumentInfo_WithIncludeTabStops_ShouldIncludeTabStops()
    {
        // Arrange
        var docPath = CreateWordDocument("test_document_info_tabs.docx");
        var arguments = CreateArguments("get_document_info", docPath);
        arguments["includeTabStops"] = true;

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }
}