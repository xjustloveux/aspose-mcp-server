using Aspose.OCR;
using Aspose.OCR.Models.PreprocessingFilters;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Handlers.Ocr.Preprocessing;

/// <summary>
///     Handler for applying auto-skew correction to images for OCR preprocessing.
///     Automatically detects and corrects image tilt/rotation.
/// </summary>
[ResultType(typeof(OcrPreprocessingResult))]
public class AutoSkewOcrPreprocessingHandler : OcrPreprocessingHandlerBase
{
    /// <inheritdoc />
    public override string Operation => "auto_skew";

    /// <summary>
    ///     Applies auto-skew correction to the input image.
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

        var filters = new PreprocessingFilter { PreprocessingFilter.AutoSkew() };

        SavePreprocessedImage(p.Path, p.OutputPath, filters);

        return CreatePreprocessingResult(p, Operation, "Auto-skew correction applied");
    }
}
