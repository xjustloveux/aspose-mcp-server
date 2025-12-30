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
        using var document = new Document(outputPath);
        Assert.NotNull(document);
        Assert.True(document.Pages.Count > 0, "Document should have pages");
    }

    [Fact]
    public async Task GetProperties_ShouldReturnAllFields()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_all_properties.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("title", result);
        Assert.Contains("author", result);
        Assert.Contains("subject", result);
        Assert.Contains("keywords", result);
        Assert.Contains("creator", result);
        Assert.Contains("producer", result);
        Assert.Contains("totalPages", result);
        Assert.Contains("isEncrypted", result);
        Assert.Contains("isLinearized", result);
    }

    [Fact]
    public async Task SetProperties_WithKeywords_ShouldSetKeywords()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_set_keywords.pdf");
        var outputPath = CreateTestFilePath("test_set_keywords_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["keywords"] = "test, pdf, keywords"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Document properties updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetProperties_WithCreatorAndProducer_ShouldAttemptToSet()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_set_creator.pdf");
        var outputPath = CreateTestFilePath("test_set_creator_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["creator"] = "Test Creator",
            ["producer"] = "Test Producer"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Document properties updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetProperties_WithAllProperties_ShouldSetAll()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_set_all.pdf");
        var outputPath = CreateTestFilePath("test_set_all_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["title"] = "Full Test",
            ["author"] = "Full Author",
            ["subject"] = "Full Subject",
            ["keywords"] = "full, test",
            ["creator"] = "Full Creator",
            ["producer"] = "Full Producer"
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Document properties updated", result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public async Task SetProperties_WithNoProperties_ShouldStillSave()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_set_empty.pdf");
        var outputPath = CreateTestFilePath("test_set_empty_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "set",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.Contains("Document properties updated", result);
        Assert.True(File.Exists(outputPath));
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
}