using Aspose.Cells;
using Aspose.Pdf.Text;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Conversion;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Conversion;

/// <summary>
///     Integration tests for ConvertDocumentTool.
///     Focuses on session management, file I/O, and format validation.
///     Detailed parameter validation tests are in Handler tests.
/// </summary>
public class ConvertDocumentToolTests : TestBase
{
    private readonly ConvertDocumentTool _tool;

    public ConvertDocumentToolTests()
    {
        _tool = new ConvertDocumentTool(SessionManager);
    }

    private string CreateWordDocument(string fileName, string content = "Test Content")
    {
        var filePath = CreateTestFilePath(fileName);
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write(content);
        doc.Save(filePath);
        return filePath;
    }

    private string CreateExcelWorkbook(string fileName, string cellValue = "Test Data")
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = cellValue;
        workbook.Save(filePath);
        return filePath;
    }

    private string CreatePowerPointPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        var slide = presentation.Slides[0];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);
        shape.TextFrame.Text = "Test Slide Content";
        presentation.Save(filePath, SlidesSaveFormat.Pptx);
        return filePath;
    }

    private string CreatePdfDocument(string fileName, string content = "Test PDF Content")
    {
        var filePath = CreateTestFilePath(fileName);
        using var pdfDoc = new Aspose.Pdf.Document();
        var page = pdfDoc.Pages.Add();
        page.Paragraphs.Add(new TextFragment(content));
        pdfDoc.Save(filePath);
        return filePath;
    }

    #region File I/O Smoke Tests

    [Fact]
    public void Convert_WordToPdf_ShouldSucceed()
    {
        var docPath = CreateWordDocument("test_word_to_pdf.docx", "Word to PDF Test Content");
        var outputPath = CreateTestFilePath("test_word_to_pdf_output.pdf");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal(outputPath, result.OutputPath);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_WordToHtml_ShouldSucceed()
    {
        var docPath = CreateWordDocument("test_word_to_html.docx", "Word to HTML Test");
        var outputPath = CreateTestFilePath("test_word_to_html_output.html");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.Equal("HTML", result.TargetFormat);
        Assert.True(File.Exists(outputPath));

        var htmlContent = File.ReadAllText(outputPath);
        Assert.Contains("<html", htmlContent, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Convert_ExcelToCsv_ShouldSucceed()
    {
        var xlsxPath = CreateExcelWorkbook("test_excel_to_csv.xlsx", "Excel Data");
        var outputPath = CreateTestFilePath("test_excel_to_csv_output.csv");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.Equal("CSV", result.TargetFormat);
        Assert.True(File.Exists(outputPath));

        var csvContent = File.ReadAllText(outputPath);
        Assert.Contains("Excel Data", csvContent);
    }

    [Fact]
    public void Convert_ExcelToPdf_ShouldSucceed()
    {
        var xlsxPath = CreateExcelWorkbook("test_excel_to_pdf.xlsx");
        var outputPath = CreateTestFilePath("test_excel_to_pdf_output.pdf");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.Equal("PDF", result.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_PowerPointToPdf_ShouldSucceed()
    {
        var pptxPath = CreatePowerPointPresentation("test_ppt_to_pdf.pptx");
        var outputPath = CreateTestFilePath("test_ppt_to_pdf_output.pdf");

        var result = _tool.Execute(pptxPath, outputPath: outputPath);

        Assert.Equal("PDF", result.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    #endregion

    #region Operation Routing

    [Fact]
    public void Execute_WithUnsupportedInputFormat_ShouldThrowArgumentException()
    {
        var unsupportedPath = CreateTestFilePath("test_unsupported.xyz");
        File.WriteAllText(unsupportedPath, "Test content");
        var outputPath = CreateTestFilePath("test_unsupported_output.pdf");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(unsupportedPath, outputPath: outputPath));
        Assert.Contains("Unsupported input format", ex.Message);
    }

    [Fact]
    public void Execute_WithUnsupportedOutputFormat_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_unsupported_output.docx");
        var outputPath = CreateTestFilePath("test_unsupported_output.xyz");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(docPath, outputPath: outputPath));
        Assert.Contains("Unsupported output format", ex.Message);
    }

    [Fact]
    public void Execute_WithNoInputPathOrSessionId_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_no_input_output.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(outputPath: outputPath));
        Assert.Contains("Either inputPath or sessionId must be provided", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyOutputPath_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_no_output.docx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(docPath, outputPath: ""));
        Assert.Contains("outputPath is required", ex.Message);
    }

    #endregion

    #region Session Management

    [Fact]
    public void Convert_WordFromSessionToPdf_ShouldSucceed()
    {
        var docPath = CreateWordDocument("test_session_word.docx", "Session Word Content");
        var sessionId = OpenSession(docPath);
        var outputPath = CreateTestFilePath("test_session_word_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.Contains(sessionId, result.SourcePath);
        Assert.True(File.Exists(outputPath));
        Assert.Equal("PDF", result.TargetFormat);
    }

    [Fact]
    public void Convert_ExcelFromSessionToCsv_ShouldSucceed()
    {
        var xlsxPath = CreateExcelWorkbook("test_session_excel.xlsx", "Session Excel Data");
        var sessionId = OpenSession(xlsxPath);
        var outputPath = CreateTestFilePath("test_session_excel_output.csv");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.Contains(sessionId, result.SourcePath);
        Assert.True(File.Exists(outputPath));
        Assert.Equal("CSV", result.TargetFormat);
    }

    [Fact]
    public void Convert_PowerPointFromSessionToPdf_ShouldSucceed()
    {
        var pptxPath = CreatePowerPointPresentation("test_session_ppt.pptx");
        var sessionId = OpenSession(pptxPath);
        var outputPath = CreateTestFilePath("test_session_ppt_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.Contains(sessionId, result.SourcePath);
        Assert.True(File.Exists(outputPath));
        Assert.Equal("PDF", result.TargetFormat);
    }

    [Fact]
    public void Execute_WithInvalidSessionId_ShouldThrowKeyNotFoundException()
    {
        var outputPath = CreateTestFilePath("test_invalid_session_output.pdf");

        var ex = Assert.Throws<KeyNotFoundException>(() =>
            _tool.Execute(sessionId: "invalid_session", outputPath: outputPath));
        Assert.Contains("invalid_session", ex.Message);
    }

    [Fact]
    public void Execute_WithBothInputPathAndSessionId_ShouldPreferSessionId()
    {
        var fileDocPath = CreateWordDocument("test_file_doc.docx", "File Content");
        var sessionDocPath = CreateExcelWorkbook("test_session_doc.xlsx", "Session Data");
        var sessionId = OpenSession(sessionDocPath);
        var outputPath = CreateTestFilePath("test_prefer_session_output.csv");

        var result = _tool.Execute(fileDocPath, sessionId, outputPath);

        Assert.Contains(sessionId, result.SourcePath);
        Assert.Equal("Excel", result.SourceFormat);
        Assert.True(File.Exists(outputPath));

        var csvContent = File.ReadAllText(outputPath);
        Assert.Contains("Session Data", csvContent);
    }

    #endregion

    #region Unsupported Output Format Tests Per Document Type

    [Fact]
    public void Execute_ExcelToUnsupportedFormat_ThrowsArgumentException()
    {
        var xlsxPath = CreateExcelWorkbook("test_excel_unsupported.xlsx");
        var outputPath = CreateTestFilePath("test_excel_unsupported.xyz");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(xlsxPath, outputPath: outputPath));
        Assert.Contains("Unsupported output format for Excel", ex.Message);
    }

    [Fact]
    public void Execute_PowerPointToUnsupportedFormat_ThrowsArgumentException()
    {
        var pptxPath = CreatePowerPointPresentation("test_ppt_unsupported.pptx");
        var outputPath = CreateTestFilePath("test_ppt_unsupported.xyz");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(pptxPath, outputPath: outputPath));
        Assert.Contains("Unsupported output format for PowerPoint", ex.Message);
    }

    [Fact]
    public void Execute_PdfToUnsupportedFormat_ThrowsArgumentException()
    {
        var docPath = CreateWordDocument("test_pdf_source.docx");
        var pdfPath = CreateTestFilePath("test_pdf_source.pdf");
        _tool.Execute(docPath, outputPath: pdfPath);

        var outputPath = CreateTestFilePath("test_pdf_unsupported.xyz");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(pdfPath, outputPath: outputPath));
        Assert.Contains("Unsupported output format for PDF", ex.Message);
    }

    #endregion

    #region Additional Format Conversion Tests

    [Fact]
    public void Convert_WordToRtf_ShouldSucceed()
    {
        var docPath = CreateWordDocument("test_word_to_rtf.docx", "Word to RTF");
        var outputPath = CreateTestFilePath("test_word_to_rtf_output.rtf");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.Equal("RTF", result.TargetFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Convert_WordToTxt_ShouldSucceed()
    {
        var docPath = CreateWordDocument("test_word_to_txt.docx", "Word to TXT");
        var outputPath = CreateTestFilePath("test_word_to_txt_output.txt");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.Equal("TXT", result.TargetFormat);
        Assert.True(File.Exists(outputPath));
        var content = File.ReadAllText(outputPath);
        Assert.Contains("Word to TXT", content);
    }

    [SkippableFact]
    public void Convert_ExcelToHtml_ShouldSucceed()
    {
        SkipInEvaluationMode(AsposeLibraryType.Cells);
        var xlsxPath = CreateExcelWorkbook("test_excel_to_html.xlsx", "Excel HTML Data");
        var outputPath = CreateTestFilePath("test_excel_to_html_output.html");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.Equal("HTML", result.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_PowerPointToHtml_ShouldSucceed()
    {
        var pptxPath = CreatePowerPointPresentation("test_ppt_to_html.pptx");
        var outputPath = CreateTestFilePath("test_ppt_to_html_output.html");

        var result = _tool.Execute(pptxPath, outputPath: outputPath);

        Assert.Equal("HTML", result.TargetFormat);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_PdfToDocx_ShouldSucceed()
    {
        var docPath = CreateWordDocument("test_pdf_source_docx.docx", "PDF to DOCX");
        var pdfPath = CreateTestFilePath("test_pdf_source_docx.pdf");
        _tool.Execute(docPath, outputPath: pdfPath);

        var outputPath = CreateTestFilePath("test_pdf_to_docx_output.docx");

        var result = _tool.Execute(pdfPath, outputPath: outputPath);

        Assert.Equal("DOCX", result.TargetFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    [Fact]
    public void Convert_ExcelToOds_ShouldSucceed()
    {
        var xlsxPath = CreateExcelWorkbook("test_excel_to_ods.xlsx", "Excel ODS Data");
        var outputPath = CreateTestFilePath("test_excel_to_ods_output.ods");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.Equal("ODS", result.TargetFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion

    #region PDF to New Format Conversion Tests

    /// <summary>
    ///     Tests that PDF can be converted to EPUB format.
    /// </summary>
    [Fact]
    public void Convert_PdfToEpub_ShouldSucceed()
    {
        var pdfPath = CreatePdfDocument("test_pdf_to_epub.pdf", "PDF to EPUB content");
        var outputPath = CreateTestFilePath("test_pdf_to_epub_output.epub");

        var result = _tool.Execute(pdfPath, outputPath: outputPath);

        Assert.Equal("EPUB", result.TargetFormat);
        Assert.Equal("PDF", result.SourceFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    /// <summary>
    ///     Tests that PDF can be converted to SVG format.
    /// </summary>
    [Fact]
    public void Convert_PdfToSvg_ShouldSucceed()
    {
        var pdfPath = CreatePdfDocument("test_pdf_to_svg.pdf", "PDF to SVG content");
        var outputPath = CreateTestFilePath("test_pdf_to_svg_output.svg");

        var result = _tool.Execute(pdfPath, outputPath: outputPath);

        Assert.Equal("SVG", result.TargetFormat);
        Assert.Equal("PDF", result.SourceFormat);
        Assert.True(File.Exists(outputPath));
    }

    /// <summary>
    ///     Tests that PDF can be converted to XPS format.
    /// </summary>
    [Fact]
    public void Convert_PdfToXps_ShouldSucceed()
    {
        var pdfPath = CreatePdfDocument("test_pdf_to_xps.pdf", "PDF to XPS content");
        var outputPath = CreateTestFilePath("test_pdf_to_xps_output.xps");

        var result = _tool.Execute(pdfPath, outputPath: outputPath);

        Assert.Equal("XPS", result.TargetFormat);
        Assert.Equal("PDF", result.SourceFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    /// <summary>
    ///     Tests that PDF-to-XML format is accepted and does not throw "Unsupported output format".
    ///     XML conversion requires a tagged PDF, so an Aspose PdfException is expected
    ///     for standard (non-tagged) PDF files.
    /// </summary>
    [Fact]
    public void Convert_PdfToXml_AcceptsFormatButRequiresTaggedPdf()
    {
        var pdfPath = CreatePdfDocument("test_pdf_to_xml.pdf", "PDF to XML content");
        var outputPath = CreateTestFilePath("test_pdf_to_xml_output.xml");

        var ex = Assert.ThrowsAny<Exception>(() => _tool.Execute(pdfPath, outputPath: outputPath));
        Assert.DoesNotContain("Unsupported output format", ex.Message);
    }

    /// <summary>
    ///     Tests that PDF can be converted to PNG image format.
    /// </summary>
    [Fact]
    public void Convert_PdfToPng_ShouldProduceImageFile()
    {
        var pdfPath = CreatePdfDocument("test_pdf_to_png.pdf", "PDF to PNG content");
        var outputPath = CreateTestFilePath("test_pdf_to_png_output.png");

        var result = _tool.Execute(pdfPath, outputPath: outputPath);

        Assert.Equal("PNG", result.TargetFormat);
        Assert.Equal("PDF", result.SourceFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    /// <summary>
    ///     Tests that PDF can be converted to JPEG image format.
    /// </summary>
    [Fact]
    public void Convert_PdfToJpeg_ShouldProduceImageFile()
    {
        var pdfPath = CreatePdfDocument("test_pdf_to_jpeg.pdf", "PDF to JPEG content");
        var outputPath = CreateTestFilePath("test_pdf_to_jpeg_output.jpg");

        var result = _tool.Execute(pdfPath, outputPath: outputPath);

        Assert.Equal("JPG", result.TargetFormat);
        Assert.Equal("PDF", result.SourceFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    /// <summary>
    ///     Tests that PDF can be converted to TIFF image format.
    /// </summary>
    [Fact]
    public void Convert_PdfToTiff_ShouldProduceImageFile()
    {
        var pdfPath = CreatePdfDocument("test_pdf_to_tiff.pdf", "PDF to TIFF content");
        var outputPath = CreateTestFilePath("test_pdf_to_tiff_output.tiff");

        var result = _tool.Execute(pdfPath, outputPath: outputPath);

        Assert.Equal("TIFF", result.TargetFormat);
        Assert.Equal("PDF", result.SourceFormat);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    #endregion
}
