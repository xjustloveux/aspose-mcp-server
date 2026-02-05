using Aspose.OCR;
using Aspose.OCR.Models.PreprocessingFilters;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Handlers.Ocr.Preprocessing;

/// <summary>
///     Handler for applying automatic dewarping to images for OCR preprocessing.
///     Corrects perspective distortion from photographed or scanned documents.
/// </summary>
[ResultType(typeof(OcrPreprocessingResult))]
public class DewarpOcrPreprocessingHandler : OcrPreprocessingHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "dewarp";

    /// <summary>
    ///     Applies automatic dewarping to the input image.
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

        var filters = new PreprocessingFilter { PreprocessingFilter.AutoDewarping() };

        SavePreprocessedImage(p.Path, p.OutputPath, filters);

        return CreatePreprocessingResult(p, Operation, "Automatic dewarping applied");
    }
}
