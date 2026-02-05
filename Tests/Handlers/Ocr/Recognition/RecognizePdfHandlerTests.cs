using Aspose.OCR;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Ocr.Recognition;
using AsposeMcpServer.Results.Ocr;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Ocr.Recognition;

/// <summary>
///     Unit tests for RecognizePdfHandler.
///     Tests parameter validation, format parsing, and error handling.
/// </summary>
public class RecognizePdfHandlerTests : HandlerTestBase<AsposeOcr>
{
    private readonly RecognizePdfHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_RecognizePdf()
    {
        Assert.Equal("recognize_pdf", _handler.Operation);
    }

    #endregion

    /// <summary>
    ///     Creates a simple PDF file with text content for OCR testing.
    /// </summary>
    /// <param name="text">The text content to add to the PDF.</param>
    /// <returns>The path to the created PDF file.</returns>
    private string CreatePdfWithText(string text)
    {
        var pdfPath = Path.Combine(TestDir, $"ocr_test_{Guid.NewGuid()}.pdf");
        var doc = new Document();
        var page = doc.Pages.Add();
        page.Paragraphs.Add(new TextFragment(text));
        doc.Save(pdfPath);
        doc.Dispose();
        return pdfPath;
    }

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ShouldThrowArgumentException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "outputPath", Path.Combine(TestDir, "output.docx") },
            { "targetFormat", "docx" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("path", ex.Message, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var tempFile = CreateTempFile(".pdf", "dummy content");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempFile },
            { "targetFormat", "docx" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("outputPath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithoutTargetFormat_ShouldThrowArgumentException()
    {
        var tempFile = CreateTempFile(".pdf", "dummy content");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempFile },
            { "outputPath", Path.Combine(TestDir, "output.docx") }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("targetFormat", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNonPdfInput_ShouldThrowArgumentException()
    {
        var tempFile = CreateTempFile(".png", "dummy content");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempFile },
            { "outputPath", Path.Combine(TestDir, "output.docx") },
            { "targetFormat", "docx" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("PDF", ex.Message);
    }

    [Fact]
    public void Execute_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent.pdf") },
            { "outputPath", Path.Combine(TestDir, "output.docx") },
            { "targetFormat", "docx" }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedTargetFormat_ShouldThrowArgumentException()
    {
        var tempFile = CreateTempFile(".pdf", "dummy content");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempFile },
            { "outputPath", Path.Combine(TestDir, "output.xyz") },
            { "targetFormat", "xyz" }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("Unsupported target format", ex.Message);
    }

    #endregion

    #region Target Format Acceptance

    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pdf")]
    [InlineData("txt")]
    public void Execute_WithValidTargetFormat_ShouldNotThrowFormatException(string format)
    {
        var tempFile = CreateTempFile(".pdf", "dummy content");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempFile },
            { "outputPath", Path.Combine(TestDir, $"output.{format}") },
            { "targetFormat", format }
        });

        var ex = Record.Exception(() => _handler.Execute(context, parameters));

        if (ex is ArgumentException argEx)
            Assert.DoesNotContain("Unsupported target format", argEx.Message);
    }

    [Theory]
    [InlineData("DOCX")]
    [InlineData("Xlsx")]
    [InlineData("PDF")]
    [InlineData("TXT")]
    public void Execute_WithTargetFormatCaseInsensitive_ShouldNotThrowFormatException(string format)
    {
        var tempFile = CreateTempFile(".pdf", "dummy content");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempFile },
            { "outputPath", Path.Combine(TestDir, $"output_ci.{format.ToLower()}") },
            { "targetFormat", format }
        });

        var ex = Record.Exception(() => _handler.Execute(context, parameters));

        if (ex is ArgumentException argEx)
            Assert.DoesNotContain("Unsupported target format", argEx.Message);
    }

    #endregion

    #region Happy Path

    [SkippableFact]
    public void Execute_WithValidPdf_ShouldReturnOcrConversionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR conversion requires a valid license");
        var pdfPath = CreatePdfWithText("Hello OCR Conversion Test");
        var outputPath = Path.Combine(TestDir, "output_happy.docx");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", pdfPath },
            { "outputPath", outputPath },
            { "targetFormat", "docx" }
        });

        var result = _handler.Execute(context, parameters);

        var convResult = Assert.IsType<OcrConversionResult>(result);
        Assert.Equal(pdfPath, convResult.SourcePath);
        Assert.Equal(outputPath, convResult.OutputPath);
        Assert.Equal("docx", convResult.TargetFormat);
        Assert.True(convResult.PageCount >= 0);
        Assert.NotNull(convResult.Message);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithTxtFormat_ShouldReturnOcrConversionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR conversion requires a valid license");
        var pdfPath = CreatePdfWithText("Text format conversion test");
        var outputPath = Path.Combine(TestDir, "output_txt.txt");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", pdfPath },
            { "outputPath", outputPath },
            { "targetFormat", "txt" }
        });

        var result = _handler.Execute(context, parameters);

        var convResult = Assert.IsType<OcrConversionResult>(result);
        Assert.Equal("txt", convResult.TargetFormat);
        Assert.NotNull(convResult.Message);
        AssertNotModified(context);
    }

    #endregion
}
