using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Conversion;
using SaveFormat = Aspose.Words.SaveFormat;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Conversion;

public class ConvertToPdfToolTests : TestBase
{
    private readonly ConvertToPdfTool _tool = new();

    #region Exception Tests

    [Fact]
    public void Convert_WithUnsupportedFormat_ShouldThrowException()
    {
        var txtPath = CreateTestFilePath("test_unsupported.txt");
        File.WriteAllText(txtPath, "Test content");

        var outputPath = CreateTestFilePath("test_unsupported_output.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute(txtPath, outputPath));
    }

    #endregion

    #region General Tests

    [Fact]
    public void ConvertWordToPdf_ShouldConvertDocument()
    {
        var docPath = CreateTestFilePath("test_convert_word.pdf");
        var doc = new Document();
        doc.Save(docPath.Replace(".pdf", ".docx"));
        docPath = docPath.Replace(".pdf", ".docx");

        var outputPath = CreateTestFilePath("test_convert_word_output.pdf");
        _tool.Execute(docPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public void ConvertExcelToPdf_ShouldConvertWorkbook()
    {
        var workbookPath = CreateTestFilePath("test_convert_excel.xlsx");
        var workbook = new Workbook();
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_convert_excel_output.pdf");
        _tool.Execute(workbookPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public void ConvertPowerPointToPdf_ShouldConvertPresentation()
    {
        var pptPath = CreateTestFilePath("test_convert_ppt.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SlidesSaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_ppt_output.pdf");
        _tool.Execute(pptPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created from PowerPoint");
    }

    [Fact]
    public void ConvertRtfToPdf_ShouldConvertRtfDocument()
    {
        var rtfPath = CreateTestFilePath("test_convert_rtf.rtf");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("RTF Test Content");
        doc.Save(rtfPath, SaveFormat.Rtf);

        var outputPath = CreateTestFilePath("test_convert_rtf_output.pdf");
        _tool.Execute(rtfPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created from RTF");
    }

    [Fact]
    public void ConvertCsvToPdf_ShouldConvertCsvFile()
    {
        var csvPath = CreateTestFilePath("test_convert_csv.csv");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Name";
        workbook.Worksheets[0].Cells["B1"].Value = "Value";
        workbook.Worksheets[0].Cells["A2"].Value = "Test";
        workbook.Worksheets[0].Cells["B2"].Value = 123;
        workbook.Save(csvPath, Aspose.Cells.SaveFormat.Csv);

        var outputPath = CreateTestFilePath("test_convert_csv_output.pdf");
        _tool.Execute(csvPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created from CSV");
    }

    [Fact]
    public void ConvertOdtToPdf_ShouldConvertOdtDocument()
    {
        var odtPath = CreateTestFilePath("test_convert_odt.odt");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("ODT Test Content");
        doc.Save(odtPath, SaveFormat.Odt);

        var outputPath = CreateTestFilePath("test_convert_odt_output.pdf");
        _tool.Execute(odtPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created from ODT");
    }

    [Fact]
    public void ConvertDocToPdf_ShouldConvertOldWordFormat()
    {
        var docPath = CreateTestFilePath("test_convert_doc.doc");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("DOC Format Test");
        doc.Save(docPath, SaveFormat.Doc);

        var outputPath = CreateTestFilePath("test_convert_doc_output.pdf");
        _tool.Execute(docPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created from DOC");
    }

    [Fact]
    public void ConvertXlsToPdf_ShouldConvertOldExcelFormat()
    {
        var xlsPath = CreateTestFilePath("test_convert_xls.xls");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test Data";
        workbook.Save(xlsPath, Aspose.Cells.SaveFormat.Excel97To2003);

        var outputPath = CreateTestFilePath("test_convert_xls_output.pdf");
        _tool.Execute(xlsPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created from XLS");
    }

    #endregion

    // Note: This tool does not support session, so no Session ID Tests region
}