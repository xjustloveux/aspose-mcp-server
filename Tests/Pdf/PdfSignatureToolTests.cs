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
}