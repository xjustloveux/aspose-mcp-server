using System.ComponentModel;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Conversion;
using ModelContextProtocol.Server;
using SaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Tools.Conversion;

/// <summary>
///     Tool for converting documents between various formats with automatic source type detection.
///     Supports Word, Excel, PowerPoint, and PDF documents.
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
    /// <returns>A ConversionResult indicating the conversion result with source and output information.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when outputPath is not provided, neither inputPath nor sessionId is provided, or the input format is
    ///     unsupported.
    /// </exception>
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

Usage examples:
- Convert Word to HTML: convert_document(inputPath='doc.docx', outputPath='doc.html')
- Convert Excel to CSV: convert_document(inputPath='book.xlsx', outputPath='book.csv')
- Convert PowerPoint to PDF: convert_document(inputPath='presentation.pptx', outputPath='presentation.pdf')
- Convert from session: convert_document(sessionId='sess_xxx', outputPath='doc.pdf')")]
    public ConversionResult Execute(
        [Description("Input file path (required if no sessionId)")]
        string? inputPath = null,
        [Description("Session ID to convert document from session")]
        string? sessionId = null,
        [Description("Output file path (required, format determined by extension)")]
        string? outputPath = null)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required");

        SecurityHelper.ValidateFilePath(outputPath, nameof(outputPath), true);

        if (string.IsNullOrEmpty(inputPath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either inputPath or sessionId must be provided");

        var outputExtension = Path.GetExtension(outputPath).ToLower();

        if (!string.IsNullOrEmpty(sessionId))
            return ConvertFromSession(sessionId, outputPath, outputExtension, $"session:{sessionId}");

        SecurityHelper.ValidateFilePath(inputPath!, nameof(inputPath), true);
        return ConvertFromFile(inputPath!, outputPath, outputExtension);
    }

    /// <summary>
    ///     Converts a document from a session to the specified output format.
    /// </summary>
    /// <param name="sessionId">The session ID containing the document.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="outputExtension">The target format extension (e.g., ".pdf", ".html").</param>
    /// <param name="sourcePath">The source path for the result.</param>
    /// <returns>A ConversionResult indicating the conversion result.</returns>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    /// <exception cref="KeyNotFoundException">Thrown when the session is not found or access is denied.</exception>
    /// <exception cref="ArgumentException">Thrown when the document type is unsupported.</exception>
    private ConversionResult ConvertFromSession(string sessionId, string outputPath, string outputExtension,
        string sourcePath)
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
                var wordFormat = GetWordSaveFormat(outputExtension);
                wordDoc.Save(outputPath, wordFormat);
                sourceType = "Word";
                break;

            case DocumentType.Excel:
                var workbook = _sessionManager.GetDocument<Workbook>(sessionId, identity);
                var excelFormat = GetExcelSaveFormat(outputExtension);
                workbook.Save(outputPath, excelFormat);
                sourceType = "Excel";
                break;

            case DocumentType.PowerPoint:
                var presentation = _sessionManager.GetDocument<Presentation>(sessionId, identity);
                var pptFormat = GetPresentationSaveFormat(outputExtension);
                presentation.Save(outputPath, pptFormat);
                sourceType = "PowerPoint";
                break;

            case DocumentType.Pdf:
                var pdfDoc = _sessionManager.GetDocument<Aspose.Pdf.Document>(sessionId, identity);
                var pdfFormat = GetPdfSaveFormat(outputExtension);
                pdfDoc.Save(outputPath, pdfFormat);
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
    /// <returns>A ConversionResult indicating the conversion result.</returns>
    /// <exception cref="ArgumentException">Thrown when the input format is unsupported.</exception>
    private static ConversionResult ConvertFromFile(string inputPath, string outputPath, string outputExtension)
    {
        var inputExtension = Path.GetExtension(inputPath).ToLower();
        string sourceFormat;

        if (IsWordDocument(inputExtension))
        {
            var doc = new Document(inputPath);
            var saveFormat = GetWordSaveFormat(outputExtension);
            doc.Save(outputPath, saveFormat);
            sourceFormat = "Word";
        }
        else if (IsExcelDocument(inputExtension))
        {
            using var workbook = new Workbook(inputPath);
            var saveFormat = GetExcelSaveFormat(outputExtension);
            workbook.Save(outputPath, saveFormat);
            sourceFormat = "Excel";
        }
        else if (IsPresentationDocument(inputExtension))
        {
            using var presentation = new Presentation(inputPath);
            var saveFormat = GetPresentationSaveFormat(outputExtension);
            presentation.Save(outputPath, saveFormat);
            sourceFormat = "PowerPoint";
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

    /// <summary>
    ///     Determines whether the specified extension represents a Word document.
    /// </summary>
    /// <param name="extension">The file extension to check.</param>
    /// <returns><c>true</c> if the extension is a Word document format; otherwise, <c>false</c>.</returns>
    private static bool IsWordDocument(string extension)
    {
        return extension is ".doc" or ".docx" or ".rtf" or ".odt" or ".txt";
    }

    /// <summary>
    ///     Determines whether the specified extension represents an Excel document.
    /// </summary>
    /// <param name="extension">The file extension to check.</param>
    /// <returns><c>true</c> if the extension is an Excel document format; otherwise, <c>false</c>.</returns>
    private static bool IsExcelDocument(string extension)
    {
        return extension is ".xls" or ".xlsx" or ".csv" or ".ods";
    }

    /// <summary>
    ///     Determines whether the specified extension represents a PowerPoint presentation.
    /// </summary>
    /// <param name="extension">The file extension to check.</param>
    /// <returns><c>true</c> if the extension is a PowerPoint format; otherwise, <c>false</c>.</returns>
    private static bool IsPresentationDocument(string extension)
    {
        return extension is ".ppt" or ".pptx" or ".odp";
    }

    /// <summary>
    ///     Gets the Word save format for the specified extension.
    /// </summary>
    /// <param name="extension">The target file extension.</param>
    /// <returns>The corresponding <see cref="SaveFormat" /> value.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not supported for Word output.</exception>
    private static SaveFormat
        GetWordSaveFormat(
            string extension) // NOSONAR S1192 - File extension strings are intentionally repeated in each format converter
    {
        return extension switch
        {
            ".pdf" => SaveFormat.Pdf,
            ".docx" => SaveFormat.Docx,
            ".doc" => SaveFormat.Doc,
            ".rtf" => SaveFormat.Rtf,
            ".html" => SaveFormat.Html,
            ".txt" => SaveFormat.Text,
            ".odt" => SaveFormat.Odt,
            _ => throw new ArgumentException($"Unsupported output format for Word: {extension}")
        };
    }

    /// <summary>
    ///     Gets the Excel save format for the specified extension.
    /// </summary>
    /// <param name="extension">The target file extension.</param>
    /// <returns>The corresponding <see cref="Aspose.Cells.SaveFormat" /> value.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not supported for Excel output.</exception>
    private static Aspose.Cells.SaveFormat GetExcelSaveFormat(string extension)
    {
        return extension switch
        {
            ".pdf" => Aspose.Cells.SaveFormat.Pdf,
            ".xlsx" => Aspose.Cells.SaveFormat.Xlsx,
            ".xls" => Aspose.Cells.SaveFormat.Excel97To2003,
            ".csv" => Aspose.Cells.SaveFormat.Csv,
            ".html" => Aspose.Cells.SaveFormat.Html,
            ".ods" => Aspose.Cells.SaveFormat.Ods,
            _ => throw new ArgumentException($"Unsupported output format for Excel: {extension}")
        };
    }

    /// <summary>
    ///     Gets the PowerPoint save format for the specified extension.
    /// </summary>
    /// <param name="extension">The target file extension.</param>
    /// <returns>The corresponding <see cref="Aspose.Slides.Export.SaveFormat" /> value.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not supported for PowerPoint output.</exception>
    private static Aspose.Slides.Export.SaveFormat GetPresentationSaveFormat(string extension)
    {
        return extension switch
        {
            ".pdf" => Aspose.Slides.Export.SaveFormat.Pdf,
            ".pptx" => Aspose.Slides.Export.SaveFormat.Pptx,
            ".ppt" => Aspose.Slides.Export.SaveFormat.Ppt,
            ".html" => Aspose.Slides.Export.SaveFormat.Html,
            ".odp" => Aspose.Slides.Export.SaveFormat.Odp,
            _ => throw new ArgumentException($"Unsupported output format for PowerPoint: {extension}")
        };
    }

    /// <summary>
    ///     Gets the PDF save format for the specified extension.
    /// </summary>
    /// <param name="extension">The target file extension.</param>
    /// <returns>The corresponding <see cref="Aspose.Pdf.SaveFormat" /> value.</returns>
    /// <exception cref="ArgumentException">Thrown when the extension is not supported for PDF output.</exception>
    private static Aspose.Pdf.SaveFormat GetPdfSaveFormat(string extension)
    {
        return extension switch
        {
            ".docx" => Aspose.Pdf.SaveFormat.DocX,
            ".doc" => Aspose.Pdf.SaveFormat.Doc,
            ".html" => Aspose.Pdf.SaveFormat.Html,
            ".xlsx" => Aspose.Pdf.SaveFormat.Excel,
            ".pptx" => Aspose.Pdf.SaveFormat.Pptx,
            ".txt" => Aspose.Pdf.SaveFormat.TeX,
            _ => throw new ArgumentException($"Unsupported output format for PDF: {extension}")
        };
    }
}
