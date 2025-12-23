using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Pdf;

namespace AsposeMcpServer.Tests.Pdf;

public class PdfRedactToolTests : PdfTestBase
{
    private readonly PdfRedactTool _tool = new();

    private string CreateTestPdf(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var document = new Document();
        var page = document.Pages.Add();
        page.Paragraphs.Add(new TextFragment("Text to redact"));
        document.Save(filePath);
        return filePath;
    }

    [Fact]
    public async Task RedactArea_ShouldRedactArea()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_area.pdf");
        var outputPath = CreateTestFilePath("test_redact_area_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 50
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }

    [Fact]
    public async Task RedactWithColor_ShouldRedactWithColor()
    {
        // Arrange
        var pdfPath = CreateTestPdf("test_redact_color.pdf");
        var outputPath = CreateTestFilePath("test_redact_color_output.pdf");
        var arguments = new JsonObject
        {
            ["path"] = pdfPath,
            ["outputPath"] = outputPath,
            ["pageIndex"] = 1,
            ["x"] = 100,
            ["y"] = 100,
            ["width"] = 200,
            ["height"] = 50,
            ["fillColor"] = "255,0,0"
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "Output PDF should be created");
    }
}