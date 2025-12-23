using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Conversion;

namespace AsposeMcpServer.Tests.Conversion;

public class ConvertToPdfToolTests : TestBase
{
    private readonly ConvertToPdfTool _tool = new();

    [Fact]
    public async Task ConvertWordToPdf_ShouldConvertDocument()
    {
        // Arrange
        var docPath = CreateTestFilePath("test_convert_word.pdf");
        var doc = new Document();
        doc.Save(docPath.Replace(".pdf", ".docx"));
        docPath = docPath.Replace(".pdf", ".docx");

        var outputPath = CreateTestFilePath("test_convert_word_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = docPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public async Task ConvertExcelToPdf_ShouldConvertWorkbook()
    {
        // Arrange
        var workbookPath = CreateTestFilePath("test_convert_excel.xlsx");
        var workbook = new Workbook();
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_convert_excel_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }
}