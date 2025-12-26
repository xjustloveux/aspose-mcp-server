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
}