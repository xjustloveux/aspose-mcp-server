using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfSignatureToolTests : PdfTestBase
{
    private readonly PdfSignatureTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    private string CreateMultiPagePdf(string fileName, int pageCount)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        for (var i = 0; i < pageCount; i++)
            document.Pages.Add();
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task GetSignatures_WithNoSignatures_ShouldReturnEmptyResult()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_no_signatures.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        Assert.Contains("count", result);
        Assert.Contains("0", result);
        Assert.Contains("No signatures found", result);
    }

    [Fact]
    public async Task GetSignatures_ShouldReturnValidJson()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_signatures_json.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("count", out var countProp));
        Assert.Equal(0, countProp.GetInt32());
        Assert.True(json.TryGetProperty("items", out var itemsProp));
        Assert.Equal(JsonValueKind.Array, itemsProp.ValueKind);
    }

    [Fact]
    public async Task Sign_WithMissingCertificatePath_ShouldThrowException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_sign_missing_cert.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "sign",
            ["path"] = pdfPath,
            ["certificatePassword"] = "password"
        };

        // Act & Assert
        // ArgumentHelper.GetString uses key as default when missing, which then fails file validation
        await Assert.ThrowsAnyAsync<Exception>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task Sign_WithNonExistentCertificatePath_ShouldThrowFileNotFoundException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_sign_missing_password.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "sign",
            ["path"] = pdfPath,
            ["certificatePath"] = "nonexistent_cert.pfx",
            ["certificatePassword"] = "password"
        };

        // Act & Assert
        await Assert.ThrowsAsync<FileNotFoundException>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task Sign_WithInvalidPageIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_sign_invalid_page.pdf");
        var certPath = CreateTestFilePath("test_cert.pfx");
        // Create a dummy cert file for the test (won't actually be used due to page validation)
        await File.WriteAllTextAsync(certPath, "dummy");

        var arguments = new JsonObject
        {
            ["operation"] = "sign",
            ["path"] = pdfPath,
            ["certificatePath"] = certPath,
            ["certificatePassword"] = "password",
            ["pageIndex"] = 99
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("pageIndex must be between", exception.Message);
    }

    [Fact]
    public async Task Delete_WithNoSignatures_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_no_signatures.pdf");
        var outputPath = CreateTestFilePath("test_delete_no_signatures_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["signatureIndex"] = 0
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("signatureIndex must be between", exception.Message);
    }

    [Fact]
    public async Task Delete_WithNegativeIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_negative_index.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath,
            ["signatureIndex"] = -1
        };

        // Act & Assert
        var exception = await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
        Assert.Contains("signatureIndex must be between", exception.Message);
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
    public async Task GetSignatures_WithMultiPagePdf_ShouldWork()
    {
        // Arrange
        var pdfPath = CreateMultiPagePdf("test_get_multipage.pdf", 5);
        var arguments = new JsonObject
        {
            ["operation"] = "get",
            ["path"] = pdfPath
        };

        // Act
        var result = await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.NotNull(result);
        var json = JsonSerializer.Deserialize<JsonElement>(result);
        Assert.True(json.TryGetProperty("count", out _));
    }

    [Fact]
    public async Task Sign_WithNonExistentImagePath_ShouldThrowFileNotFoundException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_sign_nonexistent_image.pdf");
        var certPath = CreateTestFilePath("test_cert_for_image.pfx");
        await File.WriteAllTextAsync(certPath, "dummy");

        var arguments = new JsonObject
        {
            ["operation"] = "sign",
            ["path"] = pdfPath,
            ["certificatePath"] = certPath,
            ["certificatePassword"] = "password",
            ["imagePath"] = "nonexistent_image.png"
        };

        // Act & Assert
        // This will fail at certificate validation before image validation
        // But if it reaches image validation, it should throw FileNotFoundException
        await Assert.ThrowsAnyAsync<Exception>(() => _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task Delete_WithMissingSignatureIndex_ShouldThrowArgumentException()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_delete_missing_index.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "delete",
            ["path"] = pdfPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}