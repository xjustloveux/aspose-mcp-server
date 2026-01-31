using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Results.Ocr;
using AsposeMcpServer.Tests.Infrastructure;
using AsposeMcpServer.Tools.Ocr;

namespace AsposeMcpServer.Tests.Tools.Ocr;

/// <summary>
///     Unit tests for OcrRecognitionTool.
///     Tests tool construction, operation routing, and parameter validation.
/// </summary>
public class OcrRecognitionToolTests : TestBase
{
    private readonly OcrRecognitionTool _tool = new();

    #region Construction

    [Fact]
    public void Constructor_ShouldCreateToolWithHandlerRegistry()
    {
        var tool = new OcrRecognitionTool();

        Assert.NotNull(tool);
    }

    #endregion

    /// <summary>
    ///     Creates a simple PDF file with text content for OCR testing.
    /// </summary>
    /// <param name="text">The text content to add to the PDF.</param>
    /// <returns>The path to the created PDF file.</returns>
    private string CreatePdfWithText(string text)
    {
        var pdfPath = CreateTestFilePath($"ocr_test_{Guid.NewGuid()}.pdf");
        var doc = new Document();
        var page = doc.Pages.Add();
        page.Paragraphs.Add(new TextFragment(text));
        doc.Save(pdfPath);
        doc.Dispose();
        return pdfPath;
    }

    #region Operation Routing

    [Fact]
    public void Execute_WithUnknownOperation_ShouldThrowArgumentException()
    {
        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("unknown_operation", Path.Combine(TestDir, "test.png")));

        Assert.Contains("Unknown operation", ex.Message);
    }

    [Theory]
    [InlineData("RECOGNIZE")]
    [InlineData("Recognize")]
    [InlineData("recognize")]
    public void Execute_OperationIsCaseInsensitive_Recognize(string operation)
    {
        var tempFile = Path.Combine(TestDir, "nonexistent.png");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile));
    }

    [Theory]
    [InlineData("RECOGNIZE_PDF")]
    [InlineData("Recognize_Pdf")]
    [InlineData("recognize_pdf")]
    public void Execute_OperationIsCaseInsensitive_RecognizePdf(string operation)
    {
        var tempFile = CreateTestFilePath("nonexistent.pdf");

        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute(operation, tempFile, CreateTestFilePath("output.docx"),
                targetFormat: "docx"));
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_Recognize_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        Assert.Throws<FileNotFoundException>(() =>
            _tool.Execute("recognize", Path.Combine(TestDir, "nonexistent.png")));
    }

    [Fact]
    public void Execute_Recognize_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var tempFile = CreateTestFilePath("test.xyz");
        File.WriteAllText(tempFile, "dummy");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("recognize", tempFile));

        Assert.Contains("Unsupported file format", ex.Message);
    }

    [Fact]
    public void Execute_RecognizePdf_WithNonPdfInput_ShouldThrowArgumentException()
    {
        var tempFile = CreateTestFilePath("test.png");
        File.WriteAllText(tempFile, "dummy");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("recognize_pdf", tempFile, CreateTestFilePath("output.docx"),
                targetFormat: "docx"));

        Assert.Contains("PDF", ex.Message);
    }

    [Fact]
    public void Execute_RecognizePdf_WithUnsupportedTargetFormat_ShouldThrowArgumentException()
    {
        var tempFile = CreateTestFilePath("test.pdf");
        File.WriteAllText(tempFile, "dummy");

        var ex = Assert.Throws<ArgumentException>(() =>
            _tool.Execute("recognize_pdf", tempFile, CreateTestFilePath("output.xyz"),
                targetFormat: "xyz"));

        Assert.Contains("Unsupported target format", ex.Message);
    }

    #endregion

    #region Happy Path

    [SkippableFact]
    public void Execute_Recognize_WithValidPdf_ShouldReturnRecognitionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR recognition requires a valid license");
        var pdfPath = CreatePdfWithText("Tool-level OCR recognition test");

        var result = _tool.Execute("recognize", pdfPath);

        var data = GetResultData<OcrRecognitionResult>(result);
        Assert.NotNull(data.Text);
        Assert.True(data.PageCount >= 0);
        Assert.NotNull(data.Pages);
    }

    [SkippableFact]
    public void Execute_RecognizePdf_WithValidPdf_ShouldReturnConversionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR conversion requires a valid license");
        var pdfPath = CreatePdfWithText("Tool-level OCR conversion test");
        var outputPath = CreateTestFilePath("tool_output.docx");

        var result = _tool.Execute("recognize_pdf", pdfPath, outputPath, targetFormat: "docx");

        var data = GetResultData<OcrConversionResult>(result);
        Assert.Equal(pdfPath, data.SourcePath);
        Assert.Equal(outputPath, data.OutputPath);
        Assert.Equal("docx", data.TargetFormat);
        Assert.True(data.PageCount >= 0);
        Assert.NotNull(data.Message);
    }

    #endregion
}
