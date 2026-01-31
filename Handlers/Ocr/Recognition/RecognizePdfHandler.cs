using Aspose.OCR;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Ocr;

namespace AsposeMcpServer.Handlers.Ocr.Recognition;

/// <summary>
///     Handler for converting scanned PDF files to editable documents using OCR.
/// </summary>
[ResultType(typeof(OcrConversionResult))]
public class RecognizePdfHandler : OperationHandlerBase<AsposeOcr>
{
    /// <inheritdoc />
    public override string Operation => "recognize_pdf";

    /// <summary>
    ///     Converts a scanned PDF file to an editable document format using OCR.
    /// </summary>
    /// <param name="context">The OCR engine context.</param>
    /// <param name="parameters">
    ///     Required: path (PDF file path), outputPath (output file path), targetFormat (docx/xlsx/pdf/txt).
    ///     Optional: language (default: "English").
    /// </param>
    /// <returns>An <see cref="OcrConversionResult" /> containing conversion details.</returns>
    /// <exception cref="PlatformNotSupportedException">Thrown on unsupported platforms (Linux ARM64).</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input PDF file does not exist.</exception>
    /// <exception cref="ArgumentException">Thrown when the target format is not supported or the input is not a PDF.</exception>
    public override object Execute(OperationContext<AsposeOcr> context, OperationParameters parameters)
    {
        RecognizeHandler.ValidatePlatformSupport();

        var p = ExtractParameters(parameters);
        SecurityHelper.ValidateFilePath(p.Path, "path", true);
        SecurityHelper.ValidateFilePath(p.OutputPath, "outputPath", true);

        if (!File.Exists(p.Path))
            throw new FileNotFoundException($"PDF file not found: {p.Path}");

        var ext = Path.GetExtension(p.Path);
        if (!string.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException($"Input file must be a PDF. Got: {ext}");

        var saveFormat = ParseSaveFormat(p.TargetFormat);

        var ocr = context.Document;
        var input = new OcrInput(InputType.PDF);
        input.Add(p.Path);

        var settings = new RecognitionSettings
        {
            Language = RecognizeHandler.ParseLanguage(p.Language)
        };

        var results = ocr.Recognize(input, settings);
        AsposeOcr.SaveMultipageDocument(p.OutputPath, saveFormat, results);

        var fileInfo = new FileInfo(p.OutputPath);
        return new OcrConversionResult
        {
            SourcePath = p.Path,
            OutputPath = p.OutputPath,
            TargetFormat = p.TargetFormat.ToLowerInvariant(),
            PageCount = results.Count,
            AverageConfidence = 0,
            FileSize = fileInfo.Exists ? fileInfo.Length : null,
            Message = $"PDF converted to {p.TargetFormat} with {results.Count} page(s) recognized."
        };
    }

    /// <summary>
    ///     Parses a target format string to an Aspose.OCR SaveFormat enum value.
    /// </summary>
    /// <param name="format">The target format string.</param>
    /// <returns>The parsed SaveFormat enum value.</returns>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    private static SaveFormat ParseSaveFormat(string format)
    {
        return format.ToLowerInvariant() switch
        {
            "docx" => SaveFormat.Docx,
            "xlsx" => SaveFormat.Xlsx,
            "pdf" => SaveFormat.Pdf,
            "txt" => SaveFormat.Text,
            _ => throw new ArgumentException(
                $"Unsupported target format: {format}. Supported formats: docx, xlsx, pdf, txt.")
        };
    }

    /// <summary>
    ///     Extracts recognize PDF parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static RecognizePdfParameters ExtractParameters(OperationParameters parameters)
    {
        return new RecognizePdfParameters(
            parameters.GetRequired<string>("path"),
            parameters.GetRequired<string>("outputPath"),
            parameters.GetRequired<string>("targetFormat"),
            parameters.GetOptional("language", "Eng"));
    }

    /// <summary>
    ///     Parameters for the recognize PDF operation.
    /// </summary>
    /// <param name="Path">The input PDF file path.</param>
    /// <param name="OutputPath">The output file path.</param>
    /// <param name="TargetFormat">The target document format.</param>
    /// <param name="Language">The recognition language.</param>
    private sealed record RecognizePdfParameters(string Path, string OutputPath, string TargetFormat, string Language);
}
