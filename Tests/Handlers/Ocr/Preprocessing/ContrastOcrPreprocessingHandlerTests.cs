using Aspose.OCR;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Ocr.Preprocessing;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Tests.Handlers.Ocr.Preprocessing;

/// <summary>
///     Unit tests for ContrastOcrPreprocessingHandler.
///     Tests operation property, parameter validation, error handling, and happy path.
/// </summary>
public class ContrastOcrPreprocessingHandlerTests : OcrPreprocessingHandlerTestBase
{
    private readonly ContrastOcrPreprocessingHandler _handler = new();

    /// <inheritdoc />
    protected override IOperationHandler<AsposeOcr> Handler => _handler;

    /// <inheritdoc />
    protected override string ExpectedOperation => "contrast";

    #region Operation Property

    [Fact]
    public void Operation_Returns_Contrast()
    {
        Assert.Equal("contrast", _handler.Operation);
    }

    #endregion

    #region Happy Path

    [SkippableFact]
    public void Execute_WithValidImage_ShouldReturnPreprocessingResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTempImageFile();
        var outputPath = Path.Combine(TestDir, $"contrast_output_{Guid.NewGuid()}.bmp");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreatePreprocessingParameters(tempImage, outputPath);

        var result = _handler.Execute(context, parameters);

        AssertPreprocessingResult(result, tempImage, outputPath, "contrast");
        Assert.True(File.Exists(outputPath));
        var preprocessingResult = (OcrPreprocessingResult)result;
        Assert.True(preprocessingResult.FileSize > 0);
        AssertNotModified(context);
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
    public void Execute_WithoutOutputPath_ShouldThrowArgumentException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", Path.Combine(TestDir, "input.png") }
        });

        var ex = Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));

        Assert.Contains("outputPath", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Execute_WithNonexistentFile_ShouldThrowFileNotFoundException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreatePreprocessingParameters(
            Path.Combine(TestDir, "nonexistent.png"),
            Path.Combine(TestDir, "output.png"));

        Assert.Throws<FileNotFoundException>(() => _handler.Execute(context, parameters));
    }

    [Fact]
    public void Execute_WithInvalidPath_ShouldThrowArgumentException()
    {
        var context = CreateContext(new AsposeOcr());
        var parameters = CreatePreprocessingParameters(
            "../../../etc/passwd",
            Path.Combine(TestDir, "output.png"));

        Assert.Throws<ArgumentException>(() => _handler.Execute(context, parameters));
    }

    #endregion
}
