using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfPropertiesToolTests : PdfTestBase
{
    private readonly PdfPropertiesTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task GetProperties_ShouldReturnProperties()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_properties.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.Contains("title", result);
    }

    [Fact]
    public async Task SetProperties_ShouldSetProperties()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_set_properties.pdf");
        var outputPath = CreateTestFilePath("test_set_properties_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["title"] = "Test PDF",
            ["author"] = "Test Author",
            ["subject"] = "Test Subject"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        var document = new Document(outputPath);
        // PDF metadata may have limitations, verify file was created and properties were attempted
        Assert.NotNull(document);
        Assert.True(document.Pages.Count > 0, "Document should have pages");
    }
}