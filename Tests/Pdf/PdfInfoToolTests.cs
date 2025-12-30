using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfInfoToolTests : PdfTestBase
{
    private readonly PdfInfoTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Test PDF content"));
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task GetContent_ShouldReturnContent()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_content.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_content",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Content", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetStatistics_ShouldReturnStatistics()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_statistics.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_statistics",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("fileSizeBytes", result);
    }

    [Fact]
    public async Task GetContent_WithPageIndex_ShouldReturnSpecificPage()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_content_page.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_content",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("\"pageIndex\": 1", result);
        Assert.Contains("content", result);
    }

    [Fact]
    public async Task GetContent_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_invalid_page.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_content",
            ["path"] = pdfPath,
            ["pageIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_unknown_op.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "unknown",
            ["path"] = pdfPath
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("Unknown operation", exception.Message);
    }

    [Fact]
    public async Task GetContent_WithMaxPages_ShouldLimitExtraction()
    {
        // Arrange - Create a PDF with multiple pages
        var pdfPath = CreateTestFilePath("test_max_pages.pdf");
        var document = new Document();
        for (var i = 0; i < 5; i++)
        {
            var page = document.Pages.Add();
            page.Paragraphs.Add(new TextFragment($"Page {i + 1} content"));
        }

        document.Save(pdfPath);

        var arguments = new JsonObject
        {
            ["operation"] = "get_content",
            ["path"] = pdfPath,
            ["maxPages"] = 2
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("\"extractedPages\": 2", result);
        Assert.Contains("\"truncated\": true", result);
        Assert.Contains("\"totalPages\": 5", result);
    }

    [Fact]
    public async Task GetContent_WithoutMaxPages_ShouldUseDefault()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_default_max.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_content",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("extractedPages", result);
        Assert.Contains("\"truncated\": false", result);
    }

    [Fact]
    public async Task GetStatistics_ShouldReturnAllFields()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_all_stats.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_statistics",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("fileSizeBytes", result);
        Assert.Contains("fileSizeKb", result);
        Assert.Contains("totalPages", result);
        Assert.Contains("isEncrypted", result);
        Assert.Contains("isLinearized", result);
        Assert.Contains("bookmarks", result);
        Assert.Contains("formFields", result);
        Assert.Contains("totalAnnotations", result);
        Assert.Contains("totalParagraphs", result);
    }
}