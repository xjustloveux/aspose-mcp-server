using System.IO.Compression;
using Aspose.OCR;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Ocr;
using SaveFormat = Aspose.OCR.SaveFormat;

namespace AsposeMcpServer.Handlers.Ocr.Recognition;

/// <summary>
///     Handler for converting scanned PDF files to editable documents using OCR.
///     Uses SaveMultipageDocument for all output formats with optional DOCX validation:
///     Default: Returns SaveMultipageDocument output with advisory message for DOCX.
///     Validate: Checks DOCX for invalid XML values and repairs using Aspose.Words when available.
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
    ///     Required: path (PDF file path), outputPath (output file path),
    ///     targetFormat (docx/xlsx/pdf/txt/xml/json/html/epub/rtf/pdfnoimg).
    ///     Optional: language (default: "Eng"), validate (default: false), enableWord (for DOCX repair).
    /// </param>
    /// <returns>An <see cref="OcrConversionResult" /> containing conversion details.</returns>
    /// <exception cref="PlatformNotSupportedException">Thrown on unsupported platforms (Linux ARM64).</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input PDF file does not exist.</exception>
    /// <exception cref="ArgumentException">Thrown when the target format is not supported or the input is not a PDF.</exception>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when validation detects issues and no repair option is available.
    /// </exception>
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

        ValidateSaveFormat(p.TargetFormat);

        var ocr = context.Document;
        var input = new OcrInput(InputType.PDF);
        input.Add(p.Path);

        var settings = new RecognitionSettings
        {
            Language = RecognizeHandler.ParseLanguage(p.Language)
        };

        var results = ocr.Recognize(input, settings);

        var formatLower = p.TargetFormat.ToLowerInvariant();
        var saveFormat = ParseSaveFormat(formatLower);
        string? repairUsed = null;

        AsposeOcr.SaveMultipageDocument(p.OutputPath, saveFormat, results);

        if (p.Validate && formatLower == "docx" && HasInvalidXmlValues(p.OutputPath))
        {
            if (p.EnableWord)
            {
                RepairDocxWithAsposeWords(p.OutputPath);
                repairUsed = "Aspose.Words";
            }
            else
            {
                throw new InvalidOperationException(
                    "DOCX output contains invalid XML values that may prevent opening in Microsoft Word. " +
                    "Enable Word tools (--word) to automatically repair, or use validate=false to skip validation.");
            }
        }

        var fileInfo = new FileInfo(p.OutputPath);
        return new OcrConversionResult
        {
            SourcePath = p.Path,
            OutputPath = p.OutputPath,
            TargetFormat = formatLower,
            PageCount = results.Count,
            AverageConfidence = 0,
            FileSize = fileInfo.Exists ? fileInfo.Length : null,
            Message = BuildResultMessage(results.Count, formatLower, repairUsed, p.Validate)
        };
    }

    /// <summary>
    ///     Checks whether a DOCX file contains invalid XML values (infinity) in its document body.
    ///     SaveMultipageDocument may produce DOCX files with "∞" or "-∞" in XML attributes,
    ///     which are not valid OOXML values and prevent the file from opening in Microsoft Word.
    /// </summary>
    /// <param name="docxPath">The DOCX file path to check.</param>
    /// <returns>True if the file contains invalid XML values; false otherwise.</returns>
    internal static bool HasInvalidXmlValues(string docxPath)
    {
        try
        {
            using var zip = ZipFile.OpenRead(docxPath);
            var entry = zip.GetEntry("word/document.xml");
            if (entry == null)
                return false;

            using var stream = entry.Open();
            using var reader = new StreamReader(stream);
            var xml = reader.ReadToEnd();
            return xml.Contains("\"∞\"") || xml.Contains("\"-∞\"");
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    ///     Repairs a DOCX file by round-tripping through Aspose.Words.
    ///     Opens the file with Aspose.Words (which tolerates invalid XML values)
    ///     and re-saves it, producing clean OOXML output.
    /// </summary>
    /// <param name="docxPath">The DOCX file path to repair.</param>
    internal static void RepairDocxWithAsposeWords(string docxPath)
    {
        var doc = new Document(docxPath);
        doc.Save(docxPath, Aspose.Words.SaveFormat.Docx);
    }

    /// <summary>
    ///     Builds a human-readable result message.
    /// </summary>
    /// <param name="pageCount">The number of pages processed.</param>
    /// <param name="format">The target format.</param>
    /// <param name="repairUsed">The repair method name, or null if no repair was performed.</param>
    /// <param name="validated">Whether validation was enabled.</param>
    /// <returns>A descriptive message about the conversion result.</returns>
    private static string BuildResultMessage(int pageCount, string format, string? repairUsed, bool validated)
    {
        var message = $"PDF converted to {format} with {pageCount} page(s) recognized.";

        if (repairUsed != null)
            message += $" Output repaired using {repairUsed}.";
        else if (!validated && format == "docx")
            message += " Note: DOCX output may not open correctly in Microsoft Word." +
                       " Use validate=true to detect and repair with Word tools (--word).";

        return message;
    }

    /// <summary>
    ///     Validates that the target format string is a supported format.
    /// </summary>
    /// <param name="format">The target format string.</param>
    /// <exception cref="ArgumentException">Thrown when the format is not supported.</exception>
    private static void ValidateSaveFormat(string format)
    {
        var lower = format.ToLowerInvariant();
        if (lower is not ("docx" or "xlsx" or "pdf" or "txt" or "xml" or "json" or "html"
            or "epub" or "rtf" or "pdfnoimg"))
            throw new ArgumentException(
                $"Unsupported target format: {format}. " +
                "Supported formats: docx, xlsx, pdf, txt, xml, json, html, epub, rtf, pdfnoimg.");
    }

    /// <summary>
    ///     Parses a target format string to an Aspose.OCR SaveFormat enum value.
    /// </summary>
    /// <param name="format">The target format string (already lowercased).</param>
    /// <returns>The parsed SaveFormat enum value.</returns>
    private static SaveFormat ParseSaveFormat(string format)
    {
        return format switch
        {
            "docx" => SaveFormat.Docx,
            "xlsx" => SaveFormat.Xlsx,
            "pdf" => SaveFormat.Pdf,
            "txt" => SaveFormat.Text,
            "xml" => SaveFormat.Xml,
            "json" => SaveFormat.Json,
            "html" => SaveFormat.HTML,
            "epub" => SaveFormat.EPUB,
            "rtf" => SaveFormat.RTF,
            "pdfnoimg" => SaveFormat.PdfNoImg,
            _ => throw new ArgumentException($"Unsupported save format: {format}")
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
            parameters.GetOptional("language", "Eng"),
            parameters.GetOptional("validate", false),
            parameters.GetOptional("enableWord", false));
    }

    /// <summary>
    ///     Parameters for the recognize PDF operation.
    /// </summary>
    /// <param name="Path">The input PDF file path.</param>
    /// <param name="OutputPath">The output file path.</param>
    /// <param name="TargetFormat">The target document format.</param>
    /// <param name="Language">The recognition language.</param>
    /// <param name="Validate">Whether to validate and repair DOCX output.</param>
    /// <param name="EnableWord">Whether Word tools are enabled (for DOCX repair).</param>
    private sealed record RecognizePdfParameters(
        string Path,
        string OutputPath,
        string TargetFormat,
        string Language,
        bool Validate,
        bool EnableWord);
}
