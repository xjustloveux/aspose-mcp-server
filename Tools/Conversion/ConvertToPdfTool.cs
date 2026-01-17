using System.ComponentModel;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;
using SaveFormat = Aspose.Words.SaveFormat;

namespace AsposeMcpServer.Tools.Conversion;

/// <summary>
///     Tool for converting documents (Word, Excel, PowerPoint) to PDF format.
/// </summary>
[McpServerToolType]
public class ConvertToPdfTool
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
    ///     Initializes a new instance of the <see cref="ConvertToPdfTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public ConvertToPdfTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Converts a document (Word, Excel, PowerPoint) to PDF format.
    /// </summary>
    /// <param name="inputPath">Input file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID to convert document from session.</param>
    /// <param name="outputPath">Output PDF file path (required).</param>
    /// <returns>A message indicating the conversion result with output path information.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when outputPath is not provided, neither inputPath nor sessionId is provided,
    ///     or when attempting to convert PDF to PDF.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled but sessionId is provided.</exception>
    /// <exception cref="KeyNotFoundException">Thrown when the specified session is not found or access is denied.</exception>
    [McpServerTool(Name = "convert_to_pdf")]
    [Description(@"Convert any document (Word, Excel, PowerPoint) to PDF.

Usage examples:
- Convert Word to PDF: convert_to_pdf(inputPath='doc.docx', outputPath='doc.pdf')
- Convert Excel to PDF: convert_to_pdf(inputPath='book.xlsx', outputPath='book.pdf')
- Convert PowerPoint to PDF: convert_to_pdf(inputPath='presentation.pptx', outputPath='presentation.pdf')
- Convert from session: convert_to_pdf(sessionId='sess_xxx', outputPath='doc.pdf')")]
    public string Execute(
        [Description("Input file path (required if no sessionId)")]
        string? inputPath = null,
        [Description("Session ID to convert document from session")]
        string? sessionId = null,
        [Description("Output PDF file path (required)")]
        string? outputPath = null)
    {
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("outputPath is required");

        SecurityHelper.ValidateFilePath(outputPath, nameof(outputPath), true);

        if (string.IsNullOrEmpty(inputPath) && string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("Either inputPath or sessionId must be provided");

        if (!string.IsNullOrEmpty(sessionId))
            return ConvertFromSession(sessionId, outputPath);

        SecurityHelper.ValidateFilePath(inputPath!, nameof(inputPath), true);
        return ConvertFromFile(inputPath!, outputPath);
    }

    /// <summary>
    ///     Converts a document from a session to PDF format.
    /// </summary>
    /// <param name="sessionId">The session ID containing the document.</param>
    /// <param name="outputPath">The output PDF file path.</param>
    /// <returns>A message indicating the conversion result.</returns>
    /// <exception cref="InvalidOperationException">Thrown when session management is not enabled.</exception>
    /// <exception cref="KeyNotFoundException">Thrown when the session is not found or access is denied.</exception>
    /// <exception cref="ArgumentException">Thrown when the document type is PDF (cannot convert PDF to PDF) or unsupported.</exception>
    private string ConvertFromSession(string sessionId, string outputPath)
    {
        if (_sessionManager == null)
            throw new InvalidOperationException("Session management is not enabled");

        var identity = _identityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();
        var session = _sessionManager.TryGetSession(sessionId, identity)
                      ?? throw new KeyNotFoundException($"Session '{sessionId}' not found or access denied");

        switch (session.Type)
        {
            case DocumentType.Word:
                var wordDoc = _sessionManager.GetDocument<Document>(sessionId, identity);
                wordDoc.Save(outputPath, SaveFormat.Pdf);
                break;

            case DocumentType.Excel:
                var workbook = _sessionManager.GetDocument<Workbook>(sessionId, identity);
                workbook.Save(outputPath, Aspose.Cells.SaveFormat.Pdf);
                break;

            case DocumentType.PowerPoint:
                var presentation = _sessionManager.GetDocument<Presentation>(sessionId, identity);
                presentation.Save(outputPath, Aspose.Slides.Export.SaveFormat.Pdf);
                break;

            case DocumentType.Pdf:
                throw new ArgumentException("Cannot convert PDF to PDF. The document is already in PDF format.");

            default:
                throw new ArgumentException($"Unsupported document type: {session.Type}");
        }

        return $"Document from session {sessionId} converted to PDF. Output: {outputPath}";
    }

    /// <summary>
    ///     Converts a document file to PDF format.
    /// </summary>
    /// <param name="inputPath">The source document file path.</param>
    /// <param name="outputPath">The output PDF file path.</param>
    /// <returns>A message indicating the conversion result.</returns>
    /// <exception cref="ArgumentException">Thrown when the file format is not supported.</exception>
    private static string ConvertFromFile(string inputPath, string outputPath)
    {
        var extension = Path.GetExtension(inputPath).ToLower();

        switch (extension)
        {
            case ".doc":
            case ".docx":
            case ".rtf":
            case ".odt":
                var wordDoc = new Document(inputPath);
                wordDoc.Save(outputPath, SaveFormat.Pdf);
                break;

            case ".xls":
            case ".xlsx":
            case ".csv":
            case ".ods":
                using (var workbook = new Workbook(inputPath))
                {
                    workbook.Save(outputPath, Aspose.Cells.SaveFormat.Pdf);
                }

                break;

            case ".ppt":
            case ".pptx":
            case ".odp":
                using (var presentation = new Presentation(inputPath))
                {
                    presentation.Save(outputPath, Aspose.Slides.Export.SaveFormat.Pdf);
                }

                break;

            default:
                throw new ArgumentException($"Unsupported file format: {extension}");
        }

        return $"Document converted to PDF. Output: {outputPath}";
    }
}
