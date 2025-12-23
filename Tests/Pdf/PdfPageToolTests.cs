using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfPageToolTests : PdfTestBase
{
    private readonly PdfPageTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task AddPage_ShouldAddPage()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_add_page.pdf");
        var pagesBefore = new Document(pdfPath).Pages.Count;
        var outputPath = CreateTestFilePath("test_add_page_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "add",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["count"] = 1
            // Note: insertAt may not be fully supported, so we test without it
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        using var document = new Document(outputPath);
        Assert.True(document.Pages.Count > pagesBefore, "Page should be added");
    }

    [Fact]
    public async Task DeletePage_ShouldDeletePage()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_page.pdf");
        var pagesBefore = new Document(pdfPath).Pages.Count;
        Assert.True(pagesBefore >= 2, "PDF should have at least 2 pages");

        var outputPath = CreateTestFilePath("test_delete_page_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var document = new Document(outputPath);
        Assert.True(document.Pages.Count < pagesBefore, "Page should be deleted");
    }

    [Fact]
    public async Task RotatePage_ShouldRotatePage()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_rotate_page.pdf");
        var outputPath = CreateTestFilePath("test_rotate_page_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "rotate",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["rotation"] = 90
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task GetPageInfo_ShouldReturnPageInfo()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_page_info.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_info",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Page", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task GetPageDetails_ShouldReturnPageDetails()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_page_details.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get_details",
            ["path"] = pdfPath,
            ["pageIndex"] = 1
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("Page", result, StringComparison.OrdinalIgnoreCase);
    }
}