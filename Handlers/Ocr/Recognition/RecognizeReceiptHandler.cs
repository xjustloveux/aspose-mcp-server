using Aspose.OCR;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Handlers.Ocr.Recognition;

/// <summary>
///     Handler for recognizing text from receipt images using specialized OCR.
/// </summary>
[ResultType(typeof(OcrRecognitionResult))]
public class RecognizeReceiptHandler : OperationHandlerBase<AsposeOcr>
{
    /// <inheritdoc />
    public override string Operation => "recognize_receipt";

    /// <summary>
    ///     Recognizes text from a receipt image using specialized OCR settings.
    /// </summary>
    /// <param name="context">The OCR engine context.</param>
    /// <param name="parameters">
    ///     Required: path (receipt image file path).
    ///     Optional: language (default: "Eng").
    /// </param>
    /// <returns>An <see cref="OcrRecognitionResult" /> containing recognized text and metadata.</returns>
    /// <exception cref="PlatformNotSupportedException">Thrown on unsupported platforms (Linux ARM64).</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    public override object Execute(OperationContext<AsposeOcr> context, OperationParameters parameters)
    {
        RecognizeHandler.ValidatePlatformSupport();

        var path = parameters.GetRequired<string>("path");
        var language = parameters.GetOptional("language", "Eng");

        SecurityHelper.ValidateFilePath(path, "path", true);

        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}");

        var engine = context.Document;

        var settings = new ReceiptRecognitionSettings
        {
            Language = RecognizeHandler.ParseLanguage(language)
        };

        var input = new OcrInput(InputType.SingleImage);
        input.Add(path);

        var results = engine.RecognizeReceipt(input, settings);

        return RecognizeHandler.BuildRecognitionResult(results, false);
    }
}
