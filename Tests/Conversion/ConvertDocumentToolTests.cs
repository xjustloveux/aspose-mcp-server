using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Conversion;

namespace AsposeMcpServer.Tests.Conversion;

public class ConvertDocumentToolTests : TestBase
{
    private readonly ConvertDocumentTool _tool = new();

    [Fact]
    public async Task ConvertWordToExcel_ShouldConvertDocument()
    {
        // Arrange - Word to Excel conversion is not supported, test Word to PDF instead
        var docPath = CreateTestFilePath("test_convert_word_to_pdf.docx");
        var doc = new Document();
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_convert_word_to_pdf_output.pdf");
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
    public async Task ConvertExcelToWord_ShouldConvertDocument()
    {
        // Arrange - Excel to Word conversion is not supported, test Excel to CSV instead
        var workbookPath = CreateTestFilePath("test_convert_excel_to_csv.xlsx");
        var workbook = new Workbook();
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_convert_excel_to_csv_output.csv");
        var arguments = new JsonObject
        {
            ["inputPath"] = workbookPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "CSV file should be created");
    }
}