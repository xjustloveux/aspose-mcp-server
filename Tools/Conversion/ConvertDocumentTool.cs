using System.ComponentModel;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Conversion;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Conversion;

/// <summary>
///     Tool for converting documents between various formats with automatic source type detection.
///     Supports Word, Excel, PowerPoint, PDF, HTML, EPUB, Markdown, SVG, XPS, LaTeX, and MHT as input.
///     Word, Excel, and PDF can be converted to image formats (PNG, JPEG, TIFF, BMP, SVG) with per-page/sheet rendering.
/// </summary>
[McpServerToolType]
public class ConvertDocumentTool
{
    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ConvertDocumentTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public ConvertDocumentTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Converts a document between various formats with automatic source type detection.
    /// </summary>
    /// <param name="inputPath">Input file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID to convert document from session.</param>
    /// <param name="outputPath">Output file path (required, format determined by extension).</param>
    /// <param name="pageIndex">Optional 1-based page/sheet index for single page/sheet image output (Word/Excel/PDF only).</param>
    /// <param name="dpi">Optional resolution in DPI for image output (default: 150).</param>
    /// <param name="htmlEmbedImages">Whether to embed images as Base64 in HTML output (default: true).</param>
    /// <param name="htmlSingleFile">Whether to export as single HTML file without external resources (default: true).</param>
    /// <param name="jpegQuality">JPEG quality 1-100 for JPEG image output (default: 90).</param>
    /// <param name="csvSeparator">CSV field separator character (default: comma).</param>
    /// <param name="pdfCompliance">PDF/A compliance level for PDF output (e.g., PDFA1A, PDFA1B).</param>
    /// <param name="progress">Optional progress reporter for long-running operations.</param>
    /// <returns>A ConversionResult indicating the conversion result with source and output information.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when outputPath is not provided, neither inputPath nor sessionId is provided, or the input format is
    ///     unsupported.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of valid range for the document.</exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled but sessionId is provided.</exception>
    /// <exception cref="KeyNotFoundException">Thrown when the specified session is not found or access is denied.</exception>
    [McpServerTool(
        Name = "convert_document",
        Title = "Convert Document Between Formats",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [OutputSchema(typeof(ConversionResult))]
    [Description(@"Convert documents between various formats (auto-detect source type).
Supports Word, Excel, PowerPoint, PDF, HTML, EPUB, Markdown, SVG, XPS, LaTeX, MHT as input.

Usage examples:
- Convert Word to PDF: convert_document(inputPath='doc.docx', outputPath='doc.pdf')
- Convert Word to PNG: convert_document(inputPath='doc.docx', outputPath='page.png')
- Convert Word to PNG (specific page): convert_document(inputPath='doc.docx', outputPath='page.png', pageIndex=2)
- Convert Excel to CSV: convert_document(inputPath='book.xlsx', outputPath='book.csv')
- Convert Excel to PNG: convert_document(inputPath='book.xlsx', outputPath='sheet.png')
- Convert Excel to PNG (specific sheet): convert_document(inputPath='book.xlsx', outputPath='sheet.png', pageIndex=2)
- Convert Excel to HTML (single file): convert_document(inputPath='book.xlsx', outputPath='book.html')
- Convert PowerPoint to PDF: convert_document(inputPath='slides.pptx', outputPath='slides.pdf')
- Convert PDF to Word: convert_document(inputPath='document.pdf', outputPath='document.docx')
- Convert PDF to images: convert_document(inputPath='doc.pdf', outputPath='page.png')
- Convert PDF to PNG (specific page): convert_document(inputPath='doc.pdf', outputPath='page.png', pageIndex=2)
- Convert HTML to PDF: convert_document(inputPath='page.html', outputPath='page.pdf')
- Convert EPUB to PDF: convert_document(inputPath='book.epub', outputPath='book.pdf')
- Convert Markdown to PDF: convert_document(inputPath='doc.md', outputPath='doc.pdf')
- Convert from session: convert_document(sessionId='sess_xxx', outputPath='doc.pdf')")]
    public ConversionResult Execute(
        [Description("Input file path (required if no sessionId)")]
        string? inputPath = null,
        [Description("Session ID to convert document from session")]
        string? sessionId = null,
        [Description("Output file path (required, format determined by extension)")]
        string? outputPath = null,
        [Description("Page/sheet index for image output (1-based, omit for all pages/sheets, Word/Excel/PDF only)")]
        int? pageIndex = null,
        [Description("DPI for image output (default: 150)")]
        int dpi = 150,
        [Description("Embed images as Base64 in HTML output (default: true, HTML only)")]
        bool htmlEmbedImages = true,
        [Description("Export as single HTML file without external resources (default: true, HTML only)")]
        bool htmlSingleFile = true,
        [Description("JPEG quality 1-100 (default: 90, JPEG only)")]
        int jpegQuality = 90,
        [Description("CSV field separator (default: comma, CSV only)")]
        string csvSeparator = ",",
        [Description(
            "PDF/A compliance: PDFA1A, PDFA1B, PDFA2A, PDFA2U, PDFA4 (Word) or PDFA1A, PDFA1B (Excel), default: none, PDF output only")]
        string? pdfCompliance = null,
        IProgress<ProgressNotificationValue>? progress = null)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required");

        SecurityHelper.ValidateFilePath(outputPath, nameof(outputPath), true);

        if (string.IsNullOrEmpty(inputPath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either inputPath or sessionId must be provided");

        var outputExtension = Path.GetExtension(outputPath).ToLower();

        var options = new ConversionOptions
        {
            PageIndex = pageIndex,
            Dpi = dpi,
            HtmlEmbedImages = htmlEmbedImages,
            HtmlSingleFile = htmlSingleFile,
            JpegQuality = Math.Clamp(jpegQuality, 1, 100),
            CsvSeparator = csvSeparator,
            PdfCompliance = pdfCompliance
        };

        if (!string.IsNullOrEmpty(sessionId))
            return ConvertFromSession(sessionId, outputPath, outputExtension, $"session:{sessionId}", options,
                progress);

        SecurityHelper.ValidateFilePath(inputPath!, nameof(inputPath), true);
        return ConvertFromFile(inputPath!, outputPath, outputExtension, options, progress);
    }

    /// <summary>
    ///     Converts a document from a session to the specified output format.
    /// </summary>
    /// <param name="sessionId">The session ID containing the document.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputExtension">The target format extension (e.g., ".pdf", ".html").</param>
    /// <param name="sourcePath">The source path for the result.</param>
    /// <param name="options">Conversion options including page index, DPI, and format-specific settings.</param>
    /// <param name="progress">Optional progress reporter for long-running operations.</param>
    /// <returns>A ConversionResult indicating the conversion result.</returns>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    /// <exception cref="KeyNotFoundException">Thrown when the session is not found or access is denied.</exception>
    /// <exception cref="ArgumentException">Thrown when the document type or output format is unsupported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of valid range for the document.</exception>
    private ConversionResult ConvertFromSession(string sessionId, string outputPath, string outputExtension,
        string sourcePath, ConversionOptions options, IProgress<ProgressNotificationValue>? progress)
    {
        if (_sessionManager == null)
            throw new InvalidOperationException("Session management is not enabled");

        var identity = _identityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
        var session = _sessionManager.TryGetSession(sessionId, identity)
                      ?? throw new KeyNotFoundException($"Session '{sessionId}' not found or access denied");

        string sourceType;
        switch (session.Type)
        {
            case DocumentType.Word:
                var wordDoc = _sessionManager.GetDocument<Document>(sessionId, identity);
                DocumentConversionService.ConvertWordDocument(wordDoc, outputPath, outputExtension, progress, options);
                sourceType = "Word";
                break;

            case DocumentType.Excel:
                var workbook = _sessionManager.GetDocument<Workbook>(sessionId, identity);
                DocumentConversionService.ConvertExcelDocument(workbook, outputPath, outputExtension, progress,
                    options);
                sourceType = "Excel";
                break;

            case DocumentType.PowerPoint:
                var presentation = _sessionManager.GetDocument<Presentation>(sessionId, identity);
                DocumentConversionService.ConvertPowerPointDocument(presentation, outputPath, outputExtension,
                    progress, options);
                sourceType = "PowerPoint";
                break;

            case DocumentType.Pdf:
                var pdfDoc = _sessionManager.GetDocument<Aspose.Pdf.Document>(sessionId, identity);
                DocumentConversionService.ConvertPdfDocument(pdfDoc, outputPath, outputExtension, options);
                sourceType = "PDF";
                break;

            default:
                throw new ArgumentException($"Unsupported document type: {session.Type}");
        }

        return new ConversionResult
        {
            SourcePath = sourcePath,
            OutputPath = outputPath,
            SourceFormat = sourceType,
            TargetFormat = outputExtension.TrimStart('.').ToUpperInvariant(),
            FileSize = File.Exists(outputPath) ? new FileInfo(outputPath).Length : null,
            Message = $"Document from session {sessionId} ({sourceType}) converted to {outputExtension} format"
        };
    }

    /// <summary>
    ///     Converts a document file to the specified output format.
    /// </summary>
    /// <param name="inputPath">The source document file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputExtension">The target format extension (e.g., ".pdf", ".html").</param>
    /// <param name="options">Conversion options including page index, DPI, and format-specific settings.</param>
    /// <param name="progress">Optional progress reporter for long-running operations.</param>
    /// <returns>A ConversionResult indicating the conversion result.</returns>
    /// <exception cref="ArgumentException">Thrown when the input or output format is unsupported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when pageIndex is out of valid range for the document.</exception>
    private static ConversionResult ConvertFromFile(string inputPath, string outputPath, string outputExtension,
        ConversionOptions options, IProgress<ProgressNotificationValue>? progress)
    {
        var inputExtension = Path.GetExtension(inputPath).ToLower();
        string sourceFormat;

        if (DocumentConversionService.IsWordDocument(inputExtension))
        {
            var doc = new Document(inputPath);
            DocumentConversionService.ConvertWordDocument(doc, outputPath, outputExtension, progress, options);
            sourceFormat = "Word";
        }
        else if (DocumentConversionService.IsExcelDocument(inputExtension))
        {
            using var workbook = new Workbook(inputPath);
            DocumentConversionService.ConvertExcelDocument(workbook, outputPath, outputExtension, progress, options);
            sourceFormat = "Excel";
        }
        else if (DocumentConversionService.IsPowerPointDocument(inputExtension))
        {
            using var presentation = new Presentation(inputPath);
            DocumentConversionService.ConvertPowerPointDocument(presentation, outputPath, outputExtension, progress,
                options);
            sourceFormat = "PowerPoint";
        }
        else if (DocumentConversionService.IsPdfDocument(inputExtension))
        {
            if (DocumentConversionService.IsImageFormat(outputExtension))
            {
                DocumentConversionService.ConvertPdfToImages(inputPath, outputPath, outputExtension, options.PageIndex,
                    options);
            }
            else
            {
                using var pdfDoc = new Aspose.Pdf.Document(inputPath);
                DocumentConversionService.ConvertPdfDocument(pdfDoc, outputPath, outputExtension, options);
            }

            sourceFormat = "PDF";
        }
        else if (DocumentConversionService.IsPdfConvertibleFormat(inputExtension))
        {
            var normalizedOutput = outputExtension.TrimStart('.').ToLowerInvariant();
            if (normalizedOutput != "pdf")
                throw new ArgumentException(
                    $"Format '{inputExtension}' can only be converted to PDF, not '{outputExtension}'");

            sourceFormat = DocumentConversionService.ConvertToPdfFromSpecialFormat(inputPath, outputPath);
        }
        else
        {
            throw new ArgumentException($"Unsupported input format: {inputExtension}");
        }

        return new ConversionResult
        {
            SourcePath = inputPath,
            OutputPath = outputPath,
            SourceFormat = sourceFormat,
            TargetFormat = outputExtension.TrimStart('.').ToUpperInvariant(),
            FileSize = File.Exists(outputPath) ? new FileInfo(outputPath).Length : null,
            Message = $"Document converted from {inputExtension} to {outputExtension} format"
        };
    }
}
