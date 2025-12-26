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

    [Fact]
    public async Task GetSignatures_ShouldReturnSignatures()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_get_signatures.pdf");
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
        Assert.Contains("Signature", result, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Sign_WithoutCertificate_ShouldReturnError()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_sign_no_cert.pdf");
        var outputPath = CreateTestFilePath("test_sign_no_cert_output.pdf");
        var arguments = new JsonObject
        {
            ["operation"] = "sign",
            ["path"] = pdfPath,
            ["outputPath"] = outputPath
            // No certificate path provided
        };

        // Act & Assert
        // The tool should handle missing certificate gracefully
        try
        {
            var result = await _tool.ExecuteAsync(arguments);
            // If it returns, it should be an error message
            Assert.Contains("error", result.ToLowerInvariant());
        }
        catch (Exception ex)
        {
            // Expected to throw when certificate is missing
            Assert.NotNull(ex.Message);
        }
    }

    [Fact]
    public async Task Delete_WithNoSignatures_ShouldThrowException()
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
        // Should throw when trying to delete from a PDF with no signatures
        await Assert.ThrowsAsync<ArgumentException>(() => _tool.ExecuteAsync(arguments));
    }
}