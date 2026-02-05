using Aspose.OCR;
using Aspose.OCR.Models.PreprocessingFilters;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Handlers.Ocr.Preprocessing;

/// <summary>
///     Handler for inverting image colors for OCR preprocessing.
///     Useful for processing white-on-black or light-on-dark text images.
/// </summary>
[ResultType(typeof(OcrPreprocessingResult))]
public class InvertOcrPreprocessingHandler : OcrPreprocessingHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "invert";

    /// <summary>
    ///     Inverts the colors of the input image.
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

        var filters = new PreprocessingFilter { PreprocessingFilter.Invert() };

        SavePreprocessedImage(p.Path, p.OutputPath, filters);

        return CreatePreprocessingResult(p, Operation, "Color inversion applied");
    }
}
