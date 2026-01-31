using System.ComponentModel;
using Aspose.OCR;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Ocr;

/// <summary>
///     Tool for OCR text recognition from images and PDF files.
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Ocr.Recognition")]
[McpServerToolType]
public class OcrRecognitionTool
{
    /// <summary>
    ///     Handler registry for OCR recognition operations.
    /// </summary>
    private readonly HandlerRegistry<AsposeOcr> _handlerRegistry;

    /// <summary>
    ///     Server configuration for checking enabled tool categories.
    /// </summary>
    private readonly ServerConfig? _serverConfig;

    /// <summary>
    ///     Initializes a new instance of the <see cref="OcrRecognitionTool" /> class.
    /// </summary>
    /// <param name="serverConfig">Optional server configuration for fallback capability detection.</param>
    public OcrRecognitionTool(ServerConfig? serverConfig = null)
    {
        _serverConfig = serverConfig;
        _handlerRegistry =
            HandlerRegistry<AsposeOcr>.CreateFromNamespace("AsposeMcpServer.Handlers.Ocr.Recognition");
    }

    /// <summary>
    ///     Executes an OCR recognition operation (recognize, recognize_pdf).
    /// </summary>
    /// <param name="operation">The operation to perform: recognize, recognize_pdf.</param>
    /// <param name="path">Input file path (image or PDF).</param>
    /// <param name="outputPath">Output file path (required for recognize_pdf).</param>
    /// <param name="language">
    ///     Recognition language (default: English).
    ///     Common values: English, Chinese, Japanese, Korean, German, French, Spanish.
    /// </param>
    /// <param name="targetFormat">
    ///     Target format for recognize_pdf (docx, xlsx, pdf, txt, xml, json, html, epub, rtf,
    ///     pdfnoimg).
    /// </param>
    /// <param name="includeWords">Whether to include word-level details with bounding boxes (for recognize).</param>
    /// <param name="validate">Whether to validate and repair DOCX output (for recognize_pdf).</param>
    /// <returns>Recognition result or conversion result depending on the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    /// <exception cref="PlatformNotSupportedException">Thrown on unsupported platforms (Linux ARM64).</exception>
    [McpServerTool(
        Name = "ocr_recognition",
        Title = "OCR Text Recognition",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = true,
        UseStructuredContent = true)]
    [Description(@"Recognize text from images and PDF files using OCR. Supports 2 operations: recognize, recognize_pdf.

Usage examples:
- Recognize text from image: ocr_recognition(operation='recognize', path='image.png')
- Recognize text with language: ocr_recognition(operation='recognize', path='image.png', language='Chi')
- Recognize with word details: ocr_recognition(operation='recognize', path='image.png', includeWords=true)
- Convert scanned PDF to DOCX: ocr_recognition(operation='recognize_pdf', path='scan.pdf', outputPath='output.docx', targetFormat='docx', language='Eng')
- Convert scanned PDF with Chinese: ocr_recognition(operation='recognize_pdf', path='scan.pdf', outputPath='output.docx', targetFormat='docx', language='Chi')
- Convert with DOCX validation: ocr_recognition(operation='recognize_pdf', path='scan.pdf', outputPath='output.docx', targetFormat='docx', validate=true)

Supported image formats: PNG, JPG, BMP, TIFF, GIF
Supported output formats (recognize_pdf): docx, xlsx, pdf, txt, xml, json, html, epub, rtf, pdfnoimg

Important: Set language parameter for non-English documents (e.g., language='Chi' for Chinese).
Note: OCR requires ONNX Runtime and is not supported on Linux ARM64.")]
    public object Execute(
        [Description(@"Operation to perform.
- 'recognize': Recognize text from an image or PDF file (required params: path; optional: language, includeWords)
- 'recognize_pdf': Convert scanned PDF to editable document (required params: path, outputPath, targetFormat; optional: language)")]
        string operation,
        [Description("Input file path (image file for recognize, PDF file for recognize_pdf)")]
        string path,
        [Description("Output file path (required for recognize_pdf)")]
        string? outputPath = null,
        [Description(
            "Recognition language for both recognize and recognize_pdf (default: Eng). Values: Eng, Chi, Deu, Fra, Spa, Ita, Por, Rus, Hin, or full names like English, Chinese, German")]
        string language = "Eng",
        [Description("Target format for recognize_pdf: docx, xlsx, pdf, txt, xml, json, html, epub, rtf, pdfnoimg")]
        string? targetFormat = null,
        [Description("Include word-level details with bounding boxes (for recognize, default: false)")]
        bool includeWords = false,
        [Description(
            "Validate and repair DOCX output for recognize_pdf. When true, checks for known compatibility issues and repairs using Word tools (--word) if available (default: false)")]
        bool validate = false)
    {
        var ocr = new AsposeOcr();

        var parameters = BuildParameters(path, outputPath, language, targetFormat, includeWords, validate);
        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<AsposeOcr>
        {
            Document = ocr,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        var effectiveOutputPath = string.Equals(operation, "recognize_pdf", StringComparison.OrdinalIgnoreCase)
            ? outputPath
            : path;

        return ResultHelper.FinalizeResult((dynamic)result, effectiveOutputPath, (string?)null);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="path">The input file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="language">The recognition language.</param>
    /// <param name="targetFormat">The target format for PDF conversion.</param>
    /// <param name="includeWords">Whether to include word-level details.</param>
    /// <param name="validate">Whether to validate and repair DOCX output.</param>
    /// <returns>OperationParameters configured for the OCR operation.</returns>
    private OperationParameters BuildParameters(
        string path,
        string? outputPath,
        string language,
        string? targetFormat,
        bool includeWords,
        bool validate)
    {
        var parameters = new OperationParameters();
        parameters.Set("path", path);
        parameters.SetIfNotNull("outputPath", outputPath);
        parameters.Set("language", language);
        parameters.SetIfNotNull("targetFormat", targetFormat);
        parameters.Set("includeWords", includeWords);
        parameters.Set("validate", validate);
        parameters.Set("enableWord", _serverConfig?.EnableWord ?? false);
        return parameters;
    }
}
