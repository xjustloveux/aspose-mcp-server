using Aspose.OCR;
using Aspose.OCR.Models.PreprocessingFilters;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Handlers.Ocr.Preprocessing;

/// <summary>
///     Handler for scaling images for OCR preprocessing.
///     Enlarges or reduces images to improve text recognition accuracy.
/// </summary>
[ResultType(typeof(OcrPreprocessingResult))]
public class ScaleOcrPreprocessingHandler : OcrPreprocessingHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "scale";

    /// <summary>
    ///     Scales the input image by the specified factor.
    /// </summary>
    /// <param name="context">The OCR engine context.</param>
    /// <param name="parameters">
    ///     Required: path (input image file path), outputPath (output file path).
    ///     Optional: scaleFactor (default: 2.0).
    /// </param>
    /// <returns>An <see cref="OcrPreprocessingResult" /> containing the preprocessing details.</returns>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or paths are invalid.</exception>
    /// <exception cref="InvalidOperationException">Thrown when preprocessing produces no output.</exception>
    public override object Execute(OperationContext<AsposeOcr> context, OperationParameters parameters)
    {
        var p = ExtractCommonParameters(parameters);
        var scaleFactor = parameters.GetOptional("scaleFactor", 2.0);

        var filters = new PreprocessingFilter { PreprocessingFilter.Scale((float)scaleFactor) };

        SavePreprocessedImage(p.Path, p.OutputPath, filters);

        return CreatePreprocessingResult(p, Operation,
            $"Image scaled by factor {scaleFactor:F1}");
    }
}
