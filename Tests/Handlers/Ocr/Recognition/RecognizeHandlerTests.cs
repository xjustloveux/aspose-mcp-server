using Aspose.OCR;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Handlers.Ocr.Recognition;
using AsposeMcpServer.Results.Ocr;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Ocr.Recognition;

/// <summary>
///     Unit tests for RecognizeHandler.
///     Tests language parsing, input type detection, platform validation, and parameter handling.
/// </summary>
public class RecognizeHandlerTests : HandlerTestBase<AsposeOcr>
{
    private readonly RecognizeHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_Recognize()
    {
        Assert.Equal("recognize", _handler.Operation);
    }

    #endregion

    #region Platform Validation

    [Fact]
    public void ValidatePlatformSupport_OnCurrentPlatform_ShouldNotThrow()
    {
        var ex = Record.Exception(RecognizeHandler.ValidatePlatformSupport);

        Assert.Null(ex);
    }

    #endregion

    #region Input Format Acceptance

    [Theory]
    [InlineData(".png")]
    [InlineData(".jpg")]
    [InlineData(".jpeg")]
    [InlineData(".bmp")]
    [InlineData(".tiff")]
    [InlineData(".tif")]
    [InlineData(".gif")]
    [InlineData(".pdf")]
    public void Execute_WithSupportedFormat_ShouldNotThrowFormatException(string extension)
    {
        var tempFile = CreateTempFile(extension, "dummy content");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempFile }
        });

        var ex = Record.Exception(() => _handler.Execute(context, parameters));

        if (ex is ArgumentException argEx)
            Assert.DoesNotContain("Unsupported file format", argEx.Message);
    }

    #endregion

    #region Read-Only Verification

    [Fact]
    public void Execute_ShouldNotModifyContext()
    {
        var tempImage = CreateTempImageFile();
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempImage }
        });

        _ = Record.Exception(() => _handler.Execute(context, parameters));

        AssertNotModified(context);
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

    #region ParseLanguage Tests

    [Theory]
    [InlineData("Eng", Language.Eng)]
    [InlineData("eng", Language.Eng)]
    [InlineData("Chi", Language.Chi)]
    [InlineData("Deu", Language.Deu)]
    [InlineData("Fra", Language.Fra)]
    [InlineData("Spa", Language.Spa)]
    [InlineData("Ita", Language.Ita)]
    [InlineData("Rus", Language.Rus)]
    public void ParseLanguage_WithAbbreviation_ShouldReturnCorrectLanguage(string input, Language expected)
    {
        var result = RecognizeHandler.ParseLanguage(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("English", Language.Eng)]
    [InlineData("Chinese", Language.Chi)]
    [InlineData("German", Language.Deu)]
    [InlineData("French", Language.Fra)]
    [InlineData("Spanish", Language.Spa)]
    [InlineData("Italian", Language.Ita)]
    [InlineData("Russian", Language.Rus)]
    [InlineData("Hindi", Language.Hin)]
    public void ParseLanguage_WithCommonName_ShouldReturnCorrectLanguage(string input, Language expected)
    {
        var result = RecognizeHandler.ParseLanguage(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("ENGLISH")]
    [InlineData("english")]
    [InlineData("CHINESE")]
    [InlineData("chinese")]
    public void ParseLanguage_CaseInsensitive_ShouldWork(string input)
    {
        var result = RecognizeHandler.ParseLanguage(input);

        Assert.NotEqual(Language.None, result);
    }

    [Theory]
    [InlineData("InvalidLanguage")]
    [InlineData("")]
    [InlineData("xyz")]
    public void ParseLanguage_WithInvalidLanguage_ShouldDefaultToEng(string input)
    {
        var result = RecognizeHandler.ParseLanguage(input);

        Assert.Equal(Language.Eng, result);
    }

    #endregion

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ShouldThrowArgumentException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("path", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent_file.png") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithUnsupportedFormat_ShouldThrowArgumentException()
    {
        var tempFile = CreateTempFile(".xyz", "dummy content");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempFile }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("Unsupported file format", ex.Message);
    }

    #endregion

    #region Happy Path

    [SkippableFact]
    public void Execute_WithValidPdf_ShouldReturnOcrRecognitionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR recognition requires a valid license");
        var pdfPath = CreatePdfWithText("Hello World OCR Test");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", pdfPath }
        });

        var result = _handler.Execute(context, parameters);

        var ocrResult = Assert.IsType<OcrRecognitionResult>(result);
        Assert.NotNull(ocrResult.Text);
        Assert.True(ocrResult.PageCount >= 0);
        Assert.NotNull(ocrResult.Pages);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithIncludeWords_ShouldReturnOcrRecognitionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR recognition requires a valid license");
        var pdfPath = CreatePdfWithText("Test OCR Word Details");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", pdfPath },
            { "includeWords", true }
        });

        var result = _handler.Execute(context, parameters);

        var ocrResult = Assert.IsType<OcrRecognitionResult>(result);
        Assert.NotNull(ocrResult.Text);
        Assert.NotNull(ocrResult.Pages);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithValidImage_ShouldReturnOcrRecognitionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR recognition requires a valid license");
        var tempImage = CreateTempImageFile();
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempImage }
        });

        var result = _handler.Execute(context, parameters);

        var ocrResult = Assert.IsType<OcrRecognitionResult>(result);
        Assert.NotNull(ocrResult.Text);
        Assert.NotNull(ocrResult.Pages);
        AssertNotModified(context);
    }

    #endregion
}
