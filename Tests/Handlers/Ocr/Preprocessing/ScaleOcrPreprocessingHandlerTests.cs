using Aspose.OCR;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Handlers.Ocr.Preprocessing;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Tests.Handlers.Ocr.Preprocessing;

/// <summary>
///     Unit tests for ScaleOcrPreprocessingHandler.
///     Tests operation property, parameter validation, error handling,
///     and scaleFactor parameter handling.
/// </summary>
public class ScaleOcrPreprocessingHandlerTests : OcrPreprocessingHandlerTestBase
{
    private readonly ScaleOcrPreprocessingHandler _handler = new();

    /// <inheritdoc />
    protected override IOperationHandler<AsposeOcr> Handler => _handler;

    /// <inheritdoc />
    protected override string ExpectedOperation => "scale";

    #region Operation Property

    [Fact]
    public void Operation_Returns_Scale()
    {
        Assert.Equal("scale", _handler.Operation);
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

    #region Happy Path

    [SkippableFact]
    public void Execute_WithValidImage_ShouldReturnPreprocessingResult()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTempImageFile();
        var outputPath = Path.Combine(TestDir, $"scale_output_{Guid.NewGuid()}.bmp");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempImage },
            { "outputPath", outputPath },
            { "scaleFactor", 2.0 }
        });

        var result = _handler.Execute(context, parameters);

        AssertPreprocessingResult(result, tempImage, outputPath, "scale");
        Assert.True(File.Exists(outputPath));
        var preprocessingResult = (OcrPreprocessingResult)result;
        Assert.True(preprocessingResult.FileSize > 0);
        Assert.Contains("2.0", preprocessingResult.Message);
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithDefaultScaleFactor_ShouldUseDefaultValue()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTempImageFile();
        var outputPath = Path.Combine(TestDir, $"scale_default_output_{Guid.NewGuid()}.bmp");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreatePreprocessingParameters(tempImage, outputPath);

        var result = _handler.Execute(context, parameters);

        AssertPreprocessingResult(result, tempImage, outputPath, "scale");
        Assert.True(File.Exists(outputPath));
        AssertNotModified(context);
    }

    [SkippableFact]
    public void Execute_WithCustomScaleFactor_ShouldApplyCustomValue()
    {
        SkipInEvaluationMode(AsposeLibraryType.Ocr, "OCR preprocessing requires a valid license");
        var tempImage = CreateTempImageFile();
        var outputPath = Path.Combine(TestDir, $"scale_custom_output_{Guid.NewGuid()}.bmp");
        var context = CreateContext(new AsposeOcr());
        var parameters = CreateParameters(new Dictionary<string, object?>
        {
            { "path", tempImage },
            { "outputPath", outputPath },
            { "scaleFactor", 3.0 }
        });

        var result = _handler.Execute(context, parameters);

        AssertPreprocessingResult(result, tempImage, outputPath, "scale");
        Assert.True(File.Exists(outputPath));
        var preprocessingResult = (OcrPreprocessingResult)result;
        Assert.Contains("3.0", preprocessingResult.Message);
        AssertNotModified(context);
    }

    #endregion
}
