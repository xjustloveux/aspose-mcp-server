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
    private readonly ConvertToPdfTool _tool;

    public ConvertToPdfToolTests()
    {
        _tool = new ConvertToPdfTool(SessionManager);
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
    ///     Creates a rich Word document with multiple paragraphs and a table
    /// </summary>
    private string CreateRichWordDocument(string fileName, out List<string> expectedContents)
    {
        var filePath = CreateTestFilePath(fileName);
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        expectedContents = new List<string>();

        builder.ParagraphFormat.StyleName = "Heading 1";
        var title = "PDF Conversion Test Document";
        builder.Writeln(title);
        expectedContents.Add(title);

        builder.ParagraphFormat.StyleName = "Normal";
        var para1 = "First paragraph for PDF conversion verification.";
        builder.Writeln(para1);
        expectedContents.Add(para1);

        var para2 = "Second paragraph with numbers 12345 and symbols @#$%.";
        builder.Writeln(para2);
        expectedContents.Add(para2);

        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell A1");
        expectedContents.Add("Cell A1");

        builder.InsertCell();
        builder.Write("Cell B1");
        expectedContents.Add("Cell B1");

        builder.EndRow();
        builder.EndTable();

        doc.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates a simple Excel workbook for basic tests
    /// </summary>
    private string CreateExcelWorkbook(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test Data";
        workbook.Save(filePath);
        return filePath;
    }

    /// <summary>
    ///     Creates a rich Excel workbook with multiple cells and data
    /// </summary>
    private string CreateRichExcelWorkbook(string fileName, out Dictionary<string, string> expectedContents)
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Name = "SalesData";

        expectedContents = new Dictionary<string, string>();

        sheet.Cells["A1"].Value = "Product";
        expectedContents["A1"] = "Product";

        sheet.Cells["B1"].Value = "Quantity";
        expectedContents["B1"] = "Quantity";

        sheet.Cells["C1"].Value = "Price";
        expectedContents["C1"] = "Price";

        sheet.Cells["A2"].Value = "Laptop";
        expectedContents["A2"] = "Laptop";

        sheet.Cells["B2"].Value = 5;
        expectedContents["B2"] = "5";

        sheet.Cells["C2"].Value = 999.99;
        expectedContents["C2"] = "999.99";

        sheet.Cells["A3"].Value = "Mouse";
        expectedContents["A3"] = "Mouse";

        sheet.Cells["B3"].Value = 20;
        expectedContents["B3"] = "20";

        sheet.Cells["C3"].Value = 25.50;
        expectedContents["C3"] = "25.5";

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
    ///     Creates a rich PowerPoint presentation with multiple slides
    /// </summary>
    private string CreateRichPowerPointPresentation(string fileName, int slideCount, out List<string> expectedContents)
    {
        var filePath = CreateTestFilePath(fileName);
        using var presentation = new Presentation();
        expectedContents = new List<string>();

        for (var i = 0; i < slideCount; i++)
        {
            var slide = i == 0
                ? presentation.Slides[0]
                : presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

            var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 600, 100);
            var text = $"Slide {i + 1} Content for PDF Verification";
            shape.TextFrame.Text = text;
            expectedContents.Add(text);
        }

        presentation.Save(filePath, SlidesSaveFormat.Pptx);
        return filePath;
    }

    #region General

    [Fact]
    public void Convert_WordDocxToPdf_ShouldSucceed()
    {
        var docPath = CreateRichWordDocument("test_docx.docx", out _);
        var outputPath = CreateTestFilePath("test_docx_output.pdf");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.Contains(outputPath, result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0, "PDF should have at least one page");

        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 1000, "PDF should have substantial content");
    }

    [Fact]
    public void Convert_WordDocToPdf_ShouldSucceed()
    {
        var docPath = CreateTestFilePath("test_doc.doc");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("DOC Format Test Title");
        builder.Writeln("First paragraph of DOC content.");
        builder.Writeln("Second paragraph with numbers 12345.");
        doc.Save(docPath, SaveFormat.Doc);

        var outputPath = CreateTestFilePath("test_doc_output.pdf");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_RtfToPdf_ShouldSucceed()
    {
        var rtfPath = CreateTestFilePath("test_rtf.rtf");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("RTF Test Content Header");
        builder.Writeln("RTF body paragraph for conversion.");
        doc.Save(rtfPath, SaveFormat.Rtf);

        var outputPath = CreateTestFilePath("test_rtf_output.pdf");

        var result = _tool.Execute(rtfPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_OdtToPdf_ShouldSucceed()
    {
        var odtPath = CreateTestFilePath("test_odt.odt");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("ODT Test Content Title");
        builder.Writeln("ODT document body paragraph.");
        doc.Save(odtPath, SaveFormat.Odt);

        var outputPath = CreateTestFilePath("test_odt_output.pdf");

        var result = _tool.Execute(odtPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_ExcelXlsxToPdf_ShouldSucceed()
    {
        var xlsxPath = CreateRichExcelWorkbook("test_xlsx.xlsx", out _);
        var outputPath = CreateTestFilePath("test_xlsx_output.pdf");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0, "PDF should have at least one page");
    }

    [Fact]
    public void Convert_ExcelXlsToPdf_ShouldSucceed()
    {
        var xlsPath = CreateTestFilePath("test_xls.xls");
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "Name";
        sheet.Cells["B1"].Value = "Score";
        sheet.Cells["A2"].Value = "Alice";
        sheet.Cells["B2"].Value = 95;
        sheet.Cells["A3"].Value = "Bob";
        sheet.Cells["B3"].Value = 87;
        workbook.Save(xlsPath, Aspose.Cells.SaveFormat.Excel97To2003);

        var outputPath = CreateTestFilePath("test_xls_output.pdf");

        var result = _tool.Execute(xlsPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_CsvToPdf_ShouldSucceed()
    {
        var csvPath = CreateTestFilePath("test_csv.csv");
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "Name";
        sheet.Cells["B1"].Value = "Value";
        sheet.Cells["C1"].Value = "Category";
        sheet.Cells["A2"].Value = "Item1";
        sheet.Cells["B2"].Value = 100;
        sheet.Cells["C2"].Value = "TypeA";
        workbook.Save(csvPath, Aspose.Cells.SaveFormat.Csv);

        var outputPath = CreateTestFilePath("test_csv_output.pdf");

        var result = _tool.Execute(csvPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_OdsToPdf_ShouldSucceed()
    {
        var odsPath = CreateTestFilePath("test_ods.ods");
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Value = "ODS Header";
        sheet.Cells["A2"].Value = "ODS Data Row 1";
        sheet.Cells["A3"].Value = "ODS Data Row 2";
        workbook.Save(odsPath, Aspose.Cells.SaveFormat.Ods);

        var outputPath = CreateTestFilePath("test_ods_output.pdf");

        var result = _tool.Execute(odsPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_PowerPointPptxToPdf_ShouldSucceed()
    {
        var pptxPath = CreateRichPowerPointPresentation("test_pptx.pptx", 3, out _);
        var outputPath = CreateTestFilePath("test_pptx_output.pdf");

        var result = _tool.Execute(pptxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.Equal(3, pdfDoc.Pages.Count);
    }

    [Fact]
    public void Convert_PowerPointPptToPdf_ShouldSucceed()
    {
        var pptPath = CreateTestFilePath("test_ppt.ppt");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 500, 100);
            shape.TextFrame.Text = "PPT Format Test Content";
            presentation.Save(pptPath, SlidesSaveFormat.Ppt);
        }

        var outputPath = CreateTestFilePath("test_ppt_output.pdf");

        var result = _tool.Execute(pptPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_OdpToPdf_ShouldSucceed()
    {
        var odpPath = CreateTestFilePath("test_odp.odp");
        using (var presentation = new Presentation())
        {
            var slide = presentation.Slides[0];
            var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 500, 100);
            shape.TextFrame.Text = "ODP Format Test Content";
            presentation.Save(odpPath, SlidesSaveFormat.Odp);
        }

        var outputPath = CreateTestFilePath("test_odp_output.pdf");

        var result = _tool.Execute(odpPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    #endregion

    #region Complex Document Tests

    [Fact]
    public void Convert_ComplexWordDocumentToPdf_ShouldPreserveStructure()
    {
        var docPath = CreateRichWordDocument("test_complex_word.docx", out _);
        var outputPath = CreateTestFilePath("test_complex_word_output.pdf");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);

        var fileInfo = new FileInfo(outputPath);
        Assert.True(fileInfo.Length > 500, "Complex document PDF should have substantial content");
    }

    [Fact]
    public void Convert_MultiSheetExcelToPdf_ShouldIncludeAllData()
    {
        var xlsxPath = CreateTestFilePath("test_multi_sheet.xlsx");
        var workbook = new Workbook();

        var sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        sheet1.Cells["A1"].Value = "Sheet1 Data";

        var sheet2 = workbook.Worksheets.Add("Sheet2");
        sheet2.Cells["A1"].Value = "Sheet2 Data";

        workbook.Save(xlsxPath);
        var outputPath = CreateTestFilePath("test_multi_sheet_output.pdf");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count >= 1, "PDF should have at least one page");
    }

    [Fact]
    public void Convert_MultiSlidePresentationToPdf_ShouldPreserveAllSlides()
    {
        var pptxPath = CreateRichPowerPointPresentation("test_multi_slide.pptx", 5, out _);
        var outputPath = CreateTestFilePath("test_multi_slide_output.pdf");

        var result = _tool.Execute(pptxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.Equal(5, pdfDoc.Pages.Count);
    }

    #endregion

    #region Exception

    [Fact]
    public void Execute_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var txtPath = CreateTestFilePath("test_unsupported.txt");
        File.WriteAllText(txtPath, "Test content");
        var outputPath = CreateTestFilePath("test_unsupported_output.pdf");

        var ex = Assert.Throws<ArgumentException>(() => _tool.Execute(txtPath, outputPath: outputPath));
        Assert.Contains("Unsupported file format", ex.Message);
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
    public void Convert_WordFromSession_ShouldSucceed()
    {
        var docPath = CreateRichWordDocument("test_session_word.docx", out _);
        var sessionId = OpenSession(docPath);
        var outputPath = CreateTestFilePath("test_session_word_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0, "PDF should have at least one page");
    }

    [Fact]
    public void Convert_ExcelFromSession_ShouldSucceed()
    {
        var xlsxPath = CreateRichExcelWorkbook("test_session_excel.xlsx", out _);
        var sessionId = OpenSession(xlsxPath);
        var outputPath = CreateTestFilePath("test_session_excel_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0, "PDF should have at least one page");
    }

    [Fact]
    public void Convert_PowerPointFromSession_ShouldSucceed()
    {
        var pptxPath = CreateRichPowerPointPresentation("test_session_ppt.pptx", 3, out _);
        var sessionId = OpenSession(pptxPath);
        var outputPath = CreateTestFilePath("test_session_ppt_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
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
        var fileDocPath = CreateWordDocument("test_file_doc.docx", "File Content - Should Not Be Used");
        var sessionDocPath = CreateRichPowerPointPresentation("test_session_ppt.pptx", 4, out _);
        var sessionId = OpenSession(sessionDocPath);
        var outputPath = CreateTestFilePath("test_prefer_session_output.pdf");

        var result = _tool.Execute(fileDocPath, sessionId, outputPath);

        Assert.Contains(sessionId, result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.Equal(4, pdfDoc.Pages.Count);
    }

    #endregion
}