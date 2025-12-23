using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Conversion;
using SaveFormat = Aspose.Words.SaveFormat;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

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

    [Fact]
    public async Task ConvertPowerPointToPdf_ShouldConvertPresentation()
    {
        // Arrange
        var pptPath = CreateTestFilePath("test_convert_ppt.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SlidesSaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_ppt_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = pptPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created from PowerPoint");
    }

    [Fact]
    public async Task ConvertRtfToPdf_ShouldConvertRtfDocument()
    {
        // Arrange
        var rtfPath = CreateTestFilePath("test_convert_rtf.rtf");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("RTF Test Content");
        doc.Save(rtfPath, SaveFormat.Rtf);

        var outputPath = CreateTestFilePath("test_convert_rtf_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = rtfPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created from RTF");
    }

    [Fact]
    public async Task ConvertCsvToPdf_ShouldConvertCsvFile()
    {
        // Arrange
        var csvPath = CreateTestFilePath("test_convert_csv.csv");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Name";
        workbook.Worksheets[0].Cells["B1"].Value = "Value";
        workbook.Worksheets[0].Cells["A2"].Value = "Test";
        workbook.Worksheets[0].Cells["B2"].Value = 123;
        workbook.Save(csvPath, Aspose.Cells.SaveFormat.Csv);

        var outputPath = CreateTestFilePath("test_convert_csv_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = csvPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created from CSV");
    }

    [Fact]
    public async Task ConvertOdtToPdf_ShouldConvertOdtDocument()
    {
        // Arrange
        var odtPath = CreateTestFilePath("test_convert_odt.odt");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("ODT Test Content");
        doc.Save(odtPath, SaveFormat.Odt);

        var outputPath = CreateTestFilePath("test_convert_odt_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = odtPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created from ODT");
    }

    [Fact]
    public async Task Convert_WithUnsupportedFormat_ShouldThrowException()
    {
        // Arrange
        var txtPath = CreateTestFilePath("test_unsupported.txt");
        await File.WriteAllTextAsync(txtPath, "Test content");

        var outputPath = CreateTestFilePath("test_unsupported_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = txtPath,
            ["outputPath"] = outputPath
        };

        // Act & Assert
        await Assert.ThrowsAsync<ArgumentException>(async () => await _tool.ExecuteAsync(arguments));
    }

    [Fact]
    public async Task ConvertDocToPdf_ShouldConvertOldWordFormat()
    {
        // Arrange
        var docPath = CreateTestFilePath("test_convert_doc.doc");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("DOC Format Test");
        doc.Save(docPath, SaveFormat.Doc);

        var outputPath = CreateTestFilePath("test_convert_doc_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = docPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created from DOC");
    }

    [Fact]
    public async Task ConvertXlsToPdf_ShouldConvertOldExcelFormat()
    {
        // Arrange
        var xlsPath = CreateTestFilePath("test_convert_xls.xls");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test Data";
        workbook.Save(xlsPath, Aspose.Cells.SaveFormat.Excel97To2003);

        var outputPath = CreateTestFilePath("test_convert_xls_output.pdf");
        var arguments = new JsonObject
        {
            ["inputPath"] = xlsPath,
            ["outputPath"] = outputPath
        };

        // Act
        await _tool.ExecuteAsync(arguments);

        // Assert
        Assert.True(File.Exists(outputPath), "PDF file should be created from XLS");
    }
}