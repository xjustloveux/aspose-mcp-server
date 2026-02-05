using Aspose.OCR;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Ocr;
using AsposeMcpServer.Tests.Infrastructure;

namespace AsposeMcpServer.Tests.Handlers.Ocr.Preprocessing;

/// <summary>
///     Base class for OCR preprocessing handler tests providing shared test infrastructure.
/// </summary>
public abstract class OcrPreprocessingHandlerTestBase : HandlerTestBase<AsposeOcr>
{
    /// <summary>
    ///     Gets the handler instance under test.
    /// </summary>
    protected abstract IOperationHandler<AsposeOcr> Handler { get; }

    /// <summary>
    ///     Gets the expected operation name for the handler.
    /// </summary>
    protected abstract string ExpectedOperation { get; }

    /// <summary>
    ///     Creates operation parameters for preprocessing with path and output path.
    /// </summary>
    /// <param name="path">The input image file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <returns>The operation parameters.</returns>
    protected static OperationParameters CreatePreprocessingParameters(string path, string outputPath)
    {
        return CreateParameters(new Dictionary<string, object?>
        {
            { "path", path },
            { "outputPath", outputPath }
        });
    }

    /// <summary>
    ///     Asserts that the result is a valid <see cref="OcrPreprocessingResult" /> with expected values.
    /// </summary>
    /// <param name="result">The handler result object.</param>
    /// <param name="expectedSourcePath">The expected source path.</param>
    /// <param name="expectedOutputPath">The expected output path.</param>
    /// <param name="expectedOperation">The expected operation name.</param>
    protected static void AssertPreprocessingResult(object result, string expectedSourcePath,
        string expectedOutputPath, string expectedOperation)
    {
        var preprocessingResult = Assert.IsType<OcrPreprocessingResult>(result);
        Assert.Equal(expectedSourcePath, preprocessingResult.SourcePath);
        Assert.Equal(expectedOutputPath, preprocessingResult.OutputPath);
        Assert.Equal(expectedOperation, preprocessingResult.Operation);
        Assert.NotNull(preprocessingResult.Message);
        Assert.NotEmpty(preprocessingResult.Message);
    }
}
