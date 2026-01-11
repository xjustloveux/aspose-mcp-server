using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Tests.Helpers;
using AsposeMcpServer.Tools.Conversion;
using SlidesSaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Tests.Tools.Conversion;

/// <summary>
///     Integration tests for ConvertToPdfTool.
///     Focuses on session management, file I/O, and format validation.
///     Detailed parameter validation tests are in Handler tests.
/// </summary>
public class ConvertToPdfToolTests : TestBase
{
    private readonly ConvertToPdfTool _tool;

    public ConvertToPdfToolTests()
    {
        _tool = new ConvertToPdfTool(SessionManager);
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

    private string CreateExcelWorkbook(string fileName)
    {
        var filePath = CreateTestFilePath(fileName);
        var workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Test Data";
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

    #region File I/O Smoke Tests

    [Fact]
    public void Convert_WordToPdf_ShouldSucceed()
    {
        var docPath = CreateWordDocument("test_word_to_pdf.docx", "Word to PDF Test");
        var outputPath = CreateTestFilePath("test_word_to_pdf_output.pdf");

        var result = _tool.Execute(docPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_ExcelToPdf_ShouldSucceed()
    {
        var xlsxPath = CreateExcelWorkbook("test_excel_to_pdf.xlsx");
        var outputPath = CreateTestFilePath("test_excel_to_pdf_output.pdf");

        var result = _tool.Execute(xlsxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_PowerPointToPdf_ShouldSucceed()
    {
        var pptxPath = CreatePowerPointPresentation("test_ppt_to_pdf.pptx");
        var outputPath = CreateTestFilePath("test_ppt_to_pdf_output.pdf");

        var result = _tool.Execute(pptxPath, outputPath: outputPath);

        Assert.StartsWith("Document converted to PDF", result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    #endregion

    #region Operation Routing

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
    public void Convert_WordFromSession_ShouldSucceed()
    {
        var docPath = CreateWordDocument("test_session_word.docx", "Session Word Content");
        var sessionId = OpenSession(docPath);
        var outputPath = CreateTestFilePath("test_session_word_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_ExcelFromSession_ShouldSucceed()
    {
        var xlsxPath = CreateExcelWorkbook("test_session_excel.xlsx");
        var sessionId = OpenSession(xlsxPath);
        var outputPath = CreateTestFilePath("test_session_excel_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
        Assert.True(File.Exists(outputPath));
    }

    [Fact]
    public void Convert_PowerPointFromSession_ShouldSucceed()
    {
        var pptxPath = CreatePowerPointPresentation("test_session_ppt.pptx");
        var sessionId = OpenSession(pptxPath);
        var outputPath = CreateTestFilePath("test_session_ppt_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.StartsWith("Document from session", result);
        Assert.Contains(sessionId, result);
        Assert.True(File.Exists(outputPath));
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
        var sessionDocPath = CreatePowerPointPresentation("test_session_ppt.pptx");
        var sessionId = OpenSession(sessionDocPath);
        var outputPath = CreateTestFilePath("test_prefer_session_output.pdf");

        var result = _tool.Execute(fileDocPath, sessionId, outputPath);

        Assert.Contains(sessionId, result);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.Single(pdfDoc.Pages);
    }

    #endregion
}
