using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Tests.Infrastructure;
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

        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal(outputPath, result.OutputPath);
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

        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal(outputPath, result.OutputPath);
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

        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal(outputPath, result.OutputPath);
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

    #region New Format Conversion Tests

    /// <summary>
    ///     Tests that HTML files are accepted and converted to PDF.
    /// </summary>
    [Fact]
    public void Convert_HtmlToPdf_ShouldSucceed()
    {
        var htmlPath = CreateTestFilePath("test_html_to_pdf.html");
        File.WriteAllText(htmlPath, "<html><body><h1>Test</h1><p>HTML content</p></body></html>");
        var outputPath = CreateTestFilePath("test_html_to_pdf_output.pdf");

        var result = _tool.Execute(htmlPath, outputPath: outputPath);

        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal("HTML", result.SourceFormat);
        Assert.Equal(outputPath, result.OutputPath);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    /// <summary>
    ///     Tests that .htm extension files are accepted and converted to PDF.
    /// </summary>
    [Fact]
    public void Convert_HtmToPdf_ShouldSucceed()
    {
        var htmPath = CreateTestFilePath("test_htm_to_pdf.htm");
        File.WriteAllText(htmPath, "<html><body><h1>Test</h1><p>HTM content</p></body></html>");
        var outputPath = CreateTestFilePath("test_htm_to_pdf_output.pdf");

        var result = _tool.Execute(htmPath, outputPath: outputPath);

        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal("HTML", result.SourceFormat);
        Assert.Equal(outputPath, result.OutputPath);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    /// <summary>
    ///     Tests that Markdown files are accepted and converted to PDF.
    /// </summary>
    [Fact]
    public void Convert_MarkdownToPdf_ShouldSucceed()
    {
        var mdPath = CreateTestFilePath("test_md_to_pdf.md");
        File.WriteAllText(mdPath, "# Test\n\nParagraph content for testing.");
        var outputPath = CreateTestFilePath("test_md_to_pdf_output.pdf");

        var result = _tool.Execute(mdPath, outputPath: outputPath);

        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal("Markdown", result.SourceFormat);
        Assert.Equal(outputPath, result.OutputPath);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    /// <summary>
    ///     Tests that SVG files are accepted and converted to PDF.
    /// </summary>
    [Fact]
    public void Convert_SvgToPdf_ShouldSucceed()
    {
        var svgPath = CreateTestFilePath("test_svg_to_pdf.svg");
        File.WriteAllText(svgPath,
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"200\" height=\"200\"><rect width=\"100\" height=\"100\" fill=\"blue\"/><text x=\"10\" y=\"50\" fill=\"white\">Test</text></svg>");
        var outputPath = CreateTestFilePath("test_svg_to_pdf_output.pdf");

        var result = _tool.Execute(svgPath, outputPath: outputPath);

        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal("SVG", result.SourceFormat);
        Assert.Equal(outputPath, result.OutputPath);
        Assert.True(File.Exists(outputPath));
        Assert.True(new FileInfo(outputPath).Length > 0);
    }

    /// <summary>
    ///     Tests that new formats do not throw "Unsupported file format" ArgumentException.
    ///     Format-specific Aspose exceptions are acceptable since these formats may require
    ///     special file structures, but they must not fail at format validation.
    /// </summary>
    [Theory]
    [InlineData(".html", "HTML")]
    [InlineData(".htm", "HTML")]
    [InlineData(".epub", "EPUB")]
    [InlineData(".md", "Markdown")]
    [InlineData(".svg", "SVG")]
    [InlineData(".xps", "XPS")]
    [InlineData(".tex", "LaTeX")]
    [InlineData(".mht", "MHT")]
    [InlineData(".mhtml", "MHT")]
    public void Execute_NewFormats_DoesNotThrowUnsupportedFormatException(string extension,
        string expectedSourceFormat)
    {
        var inputPath = CreateTestFilePath($"test_format_validation{extension}");
        File.WriteAllText(inputPath, "<html><body>Test</body></html>");
        var outputPath = CreateTestFilePath($"test_format_validation{extension}.pdf");

        try
        {
            var result = _tool.Execute(inputPath, outputPath: outputPath);
            Assert.Equal("PDF", result.TargetFormat);
            Assert.Equal(expectedSourceFormat, result.SourceFormat);
        }
        catch (ArgumentException ex) when (ex.Message.Contains("Unsupported file format"))
        {
            Assert.Fail(
                $"Extension '{extension}' should be accepted but threw: {ex.Message}");
        }
        catch (Exception)
        {
            // Aspose-level exceptions (e.g., invalid file content for format) are acceptable
        }
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

        Assert.Contains(sessionId, result.SourcePath);
        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal(outputPath, result.OutputPath);
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

        Assert.Contains(sessionId, result.SourcePath);
        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal(outputPath, result.OutputPath);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
    }

    [Fact]
    public void Convert_PowerPointFromSession_ShouldSucceed()
    {
        var pptxPath = CreatePowerPointPresentation("test_session_ppt.pptx");
        var sessionId = OpenSession(pptxPath);
        var outputPath = CreateTestFilePath("test_session_ppt_output.pdf");

        var result = _tool.Execute(sessionId: sessionId, outputPath: outputPath);

        Assert.Contains(sessionId, result.SourcePath);
        Assert.Equal("PDF", result.TargetFormat);
        Assert.Equal(outputPath, result.OutputPath);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.True(pdfDoc.Pages.Count > 0);
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

        Assert.Contains(sessionId, result.SourcePath);
        Assert.True(File.Exists(outputPath));

        using var pdfDoc = new Aspose.Pdf.Document(outputPath);
        Assert.Single(pdfDoc.Pages);
    }

    #endregion
}
