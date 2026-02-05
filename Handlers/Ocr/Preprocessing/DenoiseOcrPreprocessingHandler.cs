using Aspose.OCR;
using Aspose.OCR.Models.PreprocessingFilters;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Handlers.Ocr.Preprocessing;

/// <summary>
///     Handler for applying automatic denoising to images for OCR preprocessing.
///     Removes noise and artifacts that may interfere with text recognition.
/// </summary>
[ResultType(typeof(OcrPreprocessingResult))]
public class DenoiseOcrPreprocessingHandler : OcrPreprocessingHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "denoise";

    /// <summary>
    ///     Applies automatic denoising to the input image.
    /// </summary>
    /// <param name="context">The OCR engine context.</param>
    /// <param name="parameters">
    ///     Required: path (input image file path), outputPath (output file path).
    /// </param>
    /// <returns>An <see cref="OcrPreprocessingResult" /> containing the preprocessing details.</returns>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or paths are invalid.</exception>
    /// <exception cref="InvalidOperationException">Thrown when preprocessing produces no output.</exception>
    public override object Execute(OperationContext<AsposeOcr> context, OperationParameters parameters)
    {
        var p = ExtractCommonParameters(parameters);

        var filters = new PreprocessingFilter { PreprocessingFilter.AutoDenoising() };

        SavePreprocessedImage(p.Path, p.OutputPath, filters);

        return CreatePreprocessingResult(p, Operation, "Automatic denoising applied");
    }
}
