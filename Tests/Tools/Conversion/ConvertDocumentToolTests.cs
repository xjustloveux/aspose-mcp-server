using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Conversion;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Conversion;

public class ConvertDocumentToolTests : TestBase
{
    private readonly ConvertDocumentTool _tool = new();

    #region Exception Tests

    [Fact]
    public void Convert_WithUnsupportedInputFormat_ShouldThrowException()
    {
        var unsupportedPath = CreateTestFilePath("test_unsupported.xyz");
        File.WriteAllText(unsupportedPath, "Test content");

        var outputPath = CreateTestFilePath("test_unsupported_output.pdf");
        Assert.Throws<ArgumentException>(() => _tool.Execute(unsupportedPath, outputPath));
    }

    #endregion

    #region General Tests

    [Fact]
    public void ConvertWordToPdf_ShouldConvertDocument()
    {
        var docPath = CreateTestFilePath("test_convert_word_to_pdf.docx");
        var doc = new Document();
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_convert_word_to_pdf_output.pdf");
        _tool.Execute(docPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created");
    }

    [Fact]
    public void ConvertExcelToCsv_ShouldConvertDocument()
    {
        var workbookPath = CreateTestFilePath("test_convert_excel_to_csv.xlsx");
        var workbook = new Workbook();
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_convert_excel_to_csv_output.csv");
        _tool.Execute(workbookPath, outputPath);
        Assert.True(File.Exists(outputPath), "CSV file should be created");
    }

    [Fact]
    public void ConvertWordToHtml_ShouldConvertToHtml()
    {
        var docPath = CreateTestFilePath("test_convert_word_to_html.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("HTML Test Content");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_convert_word_to_html_output.html");
        _tool.Execute(docPath, outputPath);
        Assert.True(File.Exists(outputPath), "HTML file should be created");
        var htmlContent = File.ReadAllText(outputPath);
        Assert.Contains("HTML Test Content", htmlContent);
    }

    [Fact]
    public void ConvertExcelToHtml_ShouldConvertToHtml()
    {
        var workbookPath = CreateTestFilePath("test_convert_excel_to_html.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Excel HTML Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_convert_excel_to_html_output.html");
        _tool.Execute(workbookPath, outputPath);
        Assert.True(File.Exists(outputPath), "HTML file should be created from Excel");
    }

    [Fact]
    public void ConvertPowerPointToPdf_ShouldConvertToPdf()
    {
        var pptPath = CreateTestFilePath("test_convert_ppt_to_pdf.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SlidesSaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_ppt_to_pdf_output.pdf");
        _tool.Execute(pptPath, outputPath);
        Assert.True(File.Exists(outputPath), "PDF file should be created from PowerPoint");
    }

    [Fact]
    public void ConvertPowerPointToHtml_ShouldConvertToHtml()
    {
        var pptPath = CreateTestFilePath("test_convert_ppt_to_html.pptx");
        using (var presentation = new Presentation())
        {
            presentation.Save(pptPath, SlidesSaveFormat.Pptx);
        }

        var outputPath = CreateTestFilePath("test_convert_ppt_to_html_output.html");
        _tool.Execute(pptPath, outputPath);
        Assert.True(File.Exists(outputPath), "HTML file should be created from PowerPoint");
    }

    [Fact]
    public void ConvertWordToRtf_ShouldConvertToRtf()
    {
        var docPath = CreateTestFilePath("test_convert_word_to_rtf.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("RTF Conversion Test");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_convert_word_to_rtf_output.rtf");
        _tool.Execute(docPath, outputPath);
        Assert.True(File.Exists(outputPath), "RTF file should be created");
    }

    [Fact]
    public void ConvertWordToText_ShouldConvertToText()
    {
        var docPath = CreateTestFilePath("test_convert_word_to_txt.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write("Plain Text Conversion");
        doc.Save(docPath);

        var outputPath = CreateTestFilePath("test_convert_word_to_txt_output.txt");
        _tool.Execute(docPath, outputPath);
        Assert.True(File.Exists(outputPath), "TXT file should be created");
        var content = File.ReadAllText(outputPath);
        Assert.Contains("Plain Text Conversion", content);
    }

    [Fact]
    public void ConvertExcelToOds_ShouldConvertToOds()
    {
        var workbookPath = CreateTestFilePath("test_convert_excel_to_ods.xlsx");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "ODS Test";
        workbook.Save(workbookPath);

        var outputPath = CreateTestFilePath("test_convert_excel_to_ods_output.ods");
        _tool.Execute(workbookPath, outputPath);
        Assert.True(File.Exists(outputPath), "ODS file should be created");
    }

    #endregion

    // Note: This tool does not support session, so no Session ID Tests region
}