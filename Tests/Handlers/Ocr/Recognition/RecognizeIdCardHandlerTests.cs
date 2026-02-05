using Aspose.OCR;
using AsposeMcpServer.Handlers.Ocr.Recognition;
using AsposeMcpServer.Results.Ocr;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Ocr.Recognition;

/// <summary>
///     Unit tests for RecognizeIdCardHandler.
///     Tests operation property, parameter validation, error handling, and happy path.
/// </summary>
public class RecognizeIdCardHandlerTests : HandlerTestBase<AsposeOcr>
{
    private readonly RecognizeIdCardHandler _handler = new();

    #region Operation Property

    [Fact]
    public void Operation_Returns_RecognizeIdCard()
    {
        Assert.Equal("recognize_id_card", _handler.Operation);
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

    #region Error Handling

    [Fact]
    public void Execute_WithoutPath_ShouldThrowArgumentException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateEmptyParameters();

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("path", ex.Message, StringComparison.OrdinalIgnoreCase);
        AssertNotModified(context);
    }

    [Fact]
    public void Execute_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "nonexistent_id_card.jpg") }
        });

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidPath_ShouldThrowArgumentException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", "../../../etc/passwd" }
        });

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion

    #region Happy Path

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

    [SkippableFact]
    public void Execute_WithLanguageParameter_ShouldReturnOcrRecognitionResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR recognition requires a valid license");
        var tempImage = CreateTempImageFile();
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempImage },
            { "language", "Deu" }
        });

        var result = _handler.Execute(context, parameters);

        var ocrResult = Assert.IsType<OcrRecognitionResult>(result);
        Assert.NotNull(ocrResult.Text);
        AssertNotModified(context);
    }

    #endregion
}
