using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Conversion;
using SaveFormat = Aspose.Cells.SaveFormat;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Conversion;

public class ConvertDocumentToolTests : TestBase
{
    private readonly ConvertDocumentTool _tool;

    public ConvertDocumentToolTests()
    {
        _tool = new ConvertDocumentTool(SessionManager);
    }

    /// <summary>
    ///     Creates a simple Word document with single content for basic tests
    /// </summary>
    private string CreateWordDocument(string fileName, string content = "Test Content")
    {
        var filePath = CreateTestFilePath(fileName);
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Write(content);
        doc.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates a rich Word document with multiple paragraphs, formatting, and a table
    /// </summary>
    private string CreateRichWordDocument(string fileName, out List<string> expectedContents)
    {
        var filePath = CreateTestFilePath(fileName);
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        expectedContents = new List<string>();

        builder.ParagraphFormat.StyleName = "Heading 1";
        var title = "Document Conversion Test Title";
        builder.Writeln(title);
        expectedContents.Add(title);

        builder.ParagraphFormat.StyleName = "Normal";
        var para1 = "This is the first paragraph with important content for verification.";
        builder.Writeln(para1);
        expectedContents.Add(para1);

        var para2 = "Second paragraph contains special chars: numbers 12345, symbols @#$%.";
        builder.Writeln(para2);
        expectedContents.Add(para2);

        var para3 = "Final paragraph before the table section.";
        builder.Writeln(para3);
        expectedContents.Add(para3);

        builder.StartTable();
        builder.InsertCell();
        var cell1 = "Header1";
        builder.Write(cell1);
        expectedContents.Add(cell1);

        builder.InsertCell();
        var cell2 = "Header2";
        builder.Write(cell2);
        expectedContents.Add(cell2);

        builder.EndRow();

        builder.InsertCell();
        var cell3 = "Value1";
        builder.Write(cell3);
        expectedContents.Add(cell3);

        builder.InsertCell();
        var cell4 = "Value2";
        builder.Write(cell4);
        expectedContents.Add(cell4);

        builder.EndRow();
        builder.EndTable();

        doc.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates a simple Excel workbook for basic tests
    /// </summary>
    private string CreateExcelWorkbook(string fileName, string cellValue = "Test Data")
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = cellValue;
        workbook.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates a rich Excel workbook with multiple cells, formulas, and multiple sheets
    /// </summary>
    private string CreateRichExcelWorkbook(string fileName, out Dictionary<string, string> expectedContents)
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        var sheet1 = workbook.Worksheets[0];
        sheet1.Name = "DataSheet";

        expectedContents = new Dictionary<string, string>();

        sheet1.Cells["A1"].Value = "ProductName";
        expectedContents["A1"] = "ProductName";

        sheet1.Cells["B1"].Value = "Quantity";
        expectedContents["B1"] = "Quantity";

        sheet1.Cells["C1"].Value = "Price";
        expectedContents["C1"] = "Price";

        sheet1.Cells["D1"].Value = "Total";
        expectedContents["D1"] = "Total";

        sheet1.Cells["A2"].Value = "Widget";
        expectedContents["A2"] = "Widget";

        sheet1.Cells["B2"].Value = 10;
        expectedContents["B2"] = "10";

        sheet1.Cells["C2"].Value = 25.50;
        expectedContents["C2"] = "25.5";

        sheet1.Cells["D2"].Formula = "=B2*C2";
        expectedContents["D2_formula"] = "255";

        sheet1.Cells["A3"].Value = "Gadget";
        expectedContents["A3"] = "Gadget";

        sheet1.Cells["B3"].Value = 5;
        expectedContents["B3"] = "5";

        sheet1.Cells["C3"].Value = 100.00;
        expectedContents["C3"] = "100";

        workbook.CalculateFormula();
        workbook.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates a simple PowerPoint presentation for basic tests
    /// </summary>
    private string CreatePowerPointPresentation(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        presentation.Save(filePath, SlidesSaveFormat.Pptx);
        return filePath;
    }

    /// <summary>
    ///     Creates a rich PowerPoint presentation with multiple slides and content
    /// </summary>
    private string CreateRichPowerPointPresentation(string fileName, out List<string> expectedContents)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        expectedContents = new List<string>();

        var slide1 = presentation.Slides[0];
        var titleShape = slide1.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 600, 100);
        var titleText = "Conversion Test Presentation";
        titleShape.TextFrame.Text = titleText;
        expectedContents.Add(titleText);

        var slide2 = presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        var contentShape = slide2.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 600, 300);
        var contentText = "This slide contains important content for testing conversion accuracy.";
        contentShape.TextFrame.Text = contentText;
        expectedContents.Add(contentText);

        var slide3 = presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
        var bulletShape = slide3.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 600, 300);
        var bulletText = "Bullet point content for slide three verification.";
        bulletShape.TextFrame.Text = bulletText;
        expectedContents.Add(bulletText);

        presentation.Save(filePath, SlidesSaveFormat.Pptx);
        return filePath;
    }

    #region General

    [Fact]
    public void Convert_WordToPdf_ShouldSucceed()
    {
        var docPath = CreateRichWordDocument("test_word_to_pdf.docx", out _);
        var outputPath = CreateTestFilePath("test_word_to_pdf_output.pdf");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".docx", result);
        Assert.Contains(".pdf", result);
        Assert.Contains(outputPath, result);

        Assert.True(File.Exists(outputPath), "Output PDF file should exist");
        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 0, "Output PDF file should not be empty");

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0, "PDF should have at least one page");
    }

    [Fact]
    public void Convert_WordToHtml_ShouldSucceed()
    {
        var docPath = CreateRichWordDocument("test_word_to_html.docx", out var expectedContents);
        var outputPath = CreateTestFilePath("test_word_to_html_output.html");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".docx", result);
        Assert.Contains(".html", result);
        Assert.True(File.Exists(outputPath));

        var htmlContent = File.ReadAllText(outputPath);
        foreach (var expected in expectedContents)
            Assert.Contains(expected, htmlContent, StringComparison.OrdinalIgnoreCase);

        Assert.Contains("<html", htmlContent, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<body", htmlContent, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Convert_WordToRtf_ShouldSucceed()
    {
        var testContent = "Rich Text Format Test Content with special chars: 12345 @#$%";
        var docPath = CreateWordDocument("test_word_to_rtf.docx", testContent);
        var outputPath = CreateTestFilePath("test_word_to_rtf_output.rtf");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".docx", result);
        Assert.Contains(".rtf", result);
        Assert.True(File.Exists(outputPath));

        var rtfDoc = new Document(outputPath);
        Assert.Contains(testContent, rtfDoc.GetText());
    }

    [Fact]
    public void Convert_WordToTxt_ShouldSucceed()
    {
        var docPath = CreateRichWordDocument("test_word_to_txt.docx", out var expectedContents);
        var outputPath = CreateTestFilePath("test_word_to_txt_output.txt");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".docx", result);
        Assert.Contains(".txt", result);
        Assert.True(File.Exists(outputPath));

        var txtContent = File.ReadAllText(outputPath);
        foreach (var expected in expectedContents) Assert.Contains(expected, txtContent);
    }

    [Fact]
    public void Convert_WordToOdt_ShouldSucceed()
    {
        var testContent = "ODT Format Test Content for OpenDocument";
        var docPath = CreateWordDocument("test_word_to_odt.docx", testContent);
        var outputPath = CreateTestFilePath("test_word_to_odt_output.odt");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".docx", result);
        Assert.Contains(".odt", result);
        Assert.True(File.Exists(outputPath));

        var odtDoc = new Document(outputPath);
        Assert.Contains(testContent, odtDoc.GetText());
    }

    [Fact]
    public void Convert_ExcelToPdf_ShouldSucceed()
    {
        var xlsxPath = CreateRichExcelWorkbook("test_excel_to_pdf.xlsx", out _);
        var outputPath = CreateTestFilePath("test_excel_to_pdf_output.pdf");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".xlsx", result);
        Assert.Contains(".pdf", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0, "PDF should have at least one page");
    }

    [Fact]
    public void Convert_ExcelToCsv_ShouldSucceed()
    {
        var xlsxPath = CreateRichExcelWorkbook("test_excel_to_csv.xlsx", out _);
        var outputPath = CreateTestFilePath("test_excel_to_csv_output.csv");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".xlsx", result);
        Assert.Contains(".csv", result);
        Assert.True(File.Exists(outputPath));

        var csvContent = File.ReadAllText(outputPath);
        Assert.Contains("ProductName", csvContent);
        Assert.Contains("Widget", csvContent);
        Assert.Contains("Gadget", csvContent);
    }

    [Fact]
    public void Convert_ExcelToHtml_ShouldSucceed()
    {
        var xlsxPath = CreateRichExcelWorkbook("test_excel_to_html.xlsx", out _);
        var outputPath = CreateTestFilePath("test_excel_to_html_output.html");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".xlsx", result);
        Assert.Contains(".html", result);
        Assert.True(File.Exists(outputPath));

        var htmlContent = File.ReadAllText(outputPath);
        Assert.Contains("<html", htmlContent, StringComparison.OrdinalIgnoreCase);

        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 100, "HTML output should have substantial content");
    }

    [Fact]
    public void Convert_ExcelToOds_ShouldSucceed()
    {
        var xlsxPath = CreateRichExcelWorkbook("test_excel_to_ods.xlsx", out _);
        var outputPath = CreateTestFilePath("test_excel_to_ods_output.ods");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".xlsx", result);
        Assert.Contains(".ods", result);
        Assert.True(File.Exists(outputPath));

        var odsWorkbook = new Workbook(outputPath);
        Assert.Equal("ProductName", odsWorkbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Widget", odsWorkbook.Worksheets[0].Cells["A2"].StringValue);
    }

    [Fact]
    public void Convert_PowerPointToPdf_ShouldSucceed()
    {
        var pptxPath = CreateRichPowerPointPresentation("test_ppt_to_pdf.pptx", out _);
        var outputPath = CreateTestFilePath("test_ppt_to_pdf_output.pdf");

        var result = _tool.Execute(pptxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".pptx", result);
        Assert.Contains(".pdf", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count >= 3, "PDF should have at least 3 pages for 3 slides");
    }

    [Fact]
    public void Convert_PowerPointToHtml_ShouldSucceed()
    {
        var pptxPath = CreateRichPowerPointPresentation("test_ppt_to_html.pptx", out _);
        var outputPath = CreateTestFilePath("test_ppt_to_html_output.html");

        var result = _tool.Execute(pptxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".pptx", result);
        Assert.Contains(".html", result);
        Assert.True(File.Exists(outputPath));

        var htmlContent = File.ReadAllText(outputPath);
        Assert.Contains("<html", htmlContent, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Convert_PowerPointToOdp_ShouldSucceed()
    {
        var pptxPath = CreateRichPowerPointPresentation("test_ppt_to_odp.pptx", out _);
        var outputPath = CreateTestFilePath("test_ppt_to_odp_output.odp");

        var result = _tool.Execute(pptxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".pptx", result);
        Assert.Contains(".odp", result);
        Assert.True(File.Exists(outputPath));

        using var odpPresentation = new Presentation(outputPath);
        Assert.Equal(3, odpPresentation.Slides.Count);
    }

    [Fact]
    public void Convert_OdsToXlsx_ShouldSucceed()
    {
        var odsPath = CreateTestFilePath("test_ods_to_xlsx.ods");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "ODS Header";
        workbook.Worksheets[0].Cells["A2"].Value = "ODS Data Value";
        workbook.Worksheets[0].Cells["B2"].Value = 12345;
        workbook.Save(odsPath, SaveFormat.Ods);

        var outputPath = CreateTestFilePath("test_ods_to_xlsx_output.xlsx");

        var result = _tool.Execute(odsPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".ods", result);
        Assert.Contains(".xlsx", result);
        Assert.True(File.Exists(outputPath));

        var xlsxWorkbook = new Workbook(outputPath);
        Assert.Equal("ODS Header", xlsxWorkbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("ODS Data Value", xlsxWorkbook.Worksheets[0].Cells["A2"].StringValue);
        Assert.Equal(12345, xlsxWorkbook.Worksheets[0].Cells["B2"].IntValue);
    }

    [Fact]
    public void Convert_CsvToXlsx_ShouldSucceed()
    {
        var csvPath = CreateTestFilePath("test_csv_to_xlsx.csv");
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Name";
        workbook.Worksheets[0].Cells["B1"].Value = "Age";
        workbook.Worksheets[0].Cells["A2"].Value = "Alice";
        workbook.Worksheets[0].Cells["B2"].Value = 30;
        workbook.Save(csvPath, SaveFormat.Csv);

        var outputPath = CreateTestFilePath("test_csv_to_xlsx_output.xlsx");

        var result = _tool.Execute(csvPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".csv", result);
        Assert.Contains(".xlsx", result);
        Assert.True(File.Exists(outputPath));

        var xlsxWorkbook = new Workbook(outputPath);
        Assert.Equal("Name", xlsxWorkbook.Worksheets[0].Cells["A1"].StringValue);
        Assert.Equal("Alice", xlsxWorkbook.Worksheets[0].Cells["A2"].StringValue);
    }

    [Fact]
    public void Convert_OdtToDocx_ShouldSucceed()
    {
        var odtPath = CreateTestFilePath("test_odt_to_docx.odt");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("ODT First Paragraph");
        builder.Writeln("ODT Second Paragraph with numbers 12345");
        builder.Writeln("ODT Third Paragraph");
        doc.Save(odtPath, Aspose.Words.SaveFormat.Odt);

        var outputPath = CreateTestFilePath("test_odt_to_docx_output.docx");

        var result = _tool.Execute(odtPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".odt", result);
        Assert.Contains(".docx", result);
        Assert.True(File.Exists(outputPath));

        var docxDoc = new Document(outputPath);
        var text = docxDoc.GetText();
        Assert.Contains("ODT First Paragraph", text);
        Assert.Contains("ODT Second Paragraph", text);
        Assert.Contains("12345", text);
    }

    [Fact]
    public void Convert_RtfToPdf_ShouldSucceed()
    {
        var rtfPath = CreateTestFilePath("test_rtf_to_pdf.rtf");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("RTF Document Title");
        builder.Writeln("RTF Content for PDF conversion testing");
        doc.Save(rtfPath, Aspose.Words.SaveFormat.Rtf);

        var outputPath = CreateTestFilePath("test_rtf_to_pdf_output.pdf");

        var result = _tool.Execute(rtfPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".rtf", result);
        Assert.Contains(".pdf", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_OdpToPdf_ShouldSucceed()
    {
        var odpPath = CreateTestFilePath("test_odp_to_pdf.odp");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 500, 100);
            shape.TextFrame.Text = "ODP Slide Content";
            presentation.Save(odpPath, SlidesSaveFormat.Odp);
        }

        var outputPath = CreateTestFilePath("test_odp_to_pdf_output.pdf");

        var result = _tool.Execute(odpPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".odp", result);
        Assert.Contains(".pdf", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    #endregion

    #region Complex Document Tests

    [Fact]
    public void Convert_ComplexWordDocumentWithTableToHtml_ShouldPreserveAllContent()
    {
        var docPath = CreateRichWordDocument("test_complex_word.docx", out _);
        var outputPath = CreateTestFilePath("test_complex_word_output.html");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.True(File.Exists(outputPath));

        var htmlContent = File.ReadAllText(outputPath);

        Assert.Contains("Document Conversion Test Title", htmlContent);
        Assert.Contains("first paragraph", htmlContent);
        Assert.Contains("Second paragraph", htmlContent);
        Assert.Contains("12345", htmlContent);

        Assert.Contains("Header1", htmlContent);
        Assert.Contains("Header2", htmlContent);
        Assert.Contains("Value1", htmlContent);
        Assert.Contains("Value2", htmlContent);

        Assert.Contains("<table", htmlContent, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Convert_ComplexExcelWithFormulasToHtml_ShouldProduceValidHtml()
    {
        var xlsxPath = CreateRichExcelWorkbook("test_complex_excel.xlsx", out _);
        var outputPath = CreateTestFilePath("test_complex_excel_output.html");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.Contains(".xlsx", result);
        Assert.Contains(".html", result);
        Assert.True(File.Exists(outputPath));

        var htmlContent = File.ReadAllText(outputPath);
        Assert.Contains("<html", htmlContent, StringComparison.OrdinalIgnoreCase);

        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 100, "HTML output should have substantial content");
    }

    [Fact]
    public void Convert_MultiSlidePresentation_ShouldPreserveAllSlides()
    {
        var pptxPath = CreateRichPowerPointPresentation("test_multi_slide.pptx", out _);
        var outputPath = CreateTestFilePath("test_multi_slide_output.pdf");

        var result = _tool.Execute(pptxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted from", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.Equal(3, pdfDoc.Pages.Count);
    }

    #endregion

    #region Exception

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
    public void Execute_WithEmptyOutputPath_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_no_output.docx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(docPath, outputPath: ""));
        Assert.Contains("outputPath is required", ex.Message);
    }

    [Fact]
    public void Execute_WithNullOutputPath_ShouldThrowArgumentException()
    {
        var docPath = CreateWordDocument("test_null_output.docx");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(docPath, outputPath: null));
        Assert.Contains("outputPath is required", ex.Message);
    }

    [Fact]
    public void Execute_WithNoInputPathOrSessionId_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_no_input_output.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(outputPath: outputPath));
        Assert.Contains("Either inputPath or sessionId must be provided", ex.Message);
    }

    [Fact]
    public void Execute_WithEmptyInputPath_ShouldThrowArgumentException()
    {
        var outputPath = CreateTestFilePath("test_empty_input_output.pdf");
        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute("", outputPath: outputPath));
        Assert.Contains("Either inputPath or sessionId must be provided", ex.Message);
    }

    #endregion

    #region Session

    [Fact]
    public void Convert_WordFromSessionToPdf_ShouldSucceed()
    {
        var docPath = CreateRichWordDocument("test_session_word.docx", out _);
        var sessionId = OpenSession(docPath);
        var outputPath = CreateTestFilePath("test_session_word_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
        Assert.Contains("Word", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0, "PDF should have at least one page");
    }

    [Fact]
    public void Convert_WordFromSessionToHtml_ShouldSucceed()
    {
        var docPath = CreateRichWordDocument("test_session_word_html.docx", out _);
        var sessionId = OpenSession(docPath);
        var outputPath = CreateTestFilePath("test_session_word_html_output.html");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
        Assert.True(File.Exists(outputPath));

        var htmlContent = File.ReadAllText(outputPath);
        Assert.Contains("Document Conversion Test Title", htmlContent);
        Assert.Contains("<html", htmlContent, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Convert_ExcelFromSessionToCsv_ShouldSucceed()
    {
        var xlsxPath = CreateRichExcelWorkbook("test_session_excel.xlsx", out _);
        var sessionId = OpenSession(xlsxPath);
        var outputPath = CreateTestFilePath("test_session_excel_output.csv");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
        Assert.Contains("Excel", result);
        Assert.True(File.Exists(outputPath));

        var csvContent = File.ReadAllText(outputPath);
        Assert.Contains("ProductName", csvContent);
        Assert.Contains("Widget", csvContent);
    }

    [Fact]
    public void Convert_PowerPointFromSessionToPdf_ShouldSucceed()
    {
        var pptxPath = CreateRichPowerPointPresentation("test_session_ppt.pptx", out _);
        var sessionId = OpenSession(pptxPath);
        var outputPath = CreateTestFilePath("test_session_ppt_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
        Assert.Contains("PowerPoint", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.Equal(3, pdfDoc.Pages.Count);
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
        var fileDocPath = CreateWordDocument("test_file_doc.docx", "File Content - Should Not Appear");
        var sessionDocPath = CreateRichExcelWorkbook("test_session_doc.xlsx", out _);
        var sessionId = OpenSession(sessionDocPath);
        var outputPath = CreateTestFilePath("test_prefer_session_output.csv");

        var result = _tool.Execute(fileDocPath, sessionId, outputPath);

        Assert.Contains(sessionId, result);
        Assert.Contains("Excel", result);
        Assert.True(File.Exists(outputPath));

        var csvContent = File.ReadAllText(outputPath);
        Assert.Contains("ProductName", csvContent);
        Assert.DoesNotContain("Should Not Appear", csvContent);
    }

    #endregion
}