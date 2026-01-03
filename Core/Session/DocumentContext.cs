using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Helpers;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Provides unified document access whether from file path or session
/// </summary>
public class DocumentContext<T> : IDisposable where T : class
{
    /// <summary>
    ///     Indicates whether this context owns the document and should dispose it
    /// </summary>
    private readonly bool _ownsDocument;

    /// <summary>
    ///     Session ID if this context is from a session
    /// </summary>
    private readonly string? _sessionId;

    /// <summary>
    ///     Session manager reference for session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Tracks whether this context has been disposed
    /// </summary>
    private bool _disposed;

    /// <summary>
    ///     Creates a new document context
    /// </summary>
    /// <param name="document">The document instance</param>
    /// <param name="sessionManager">Session manager (null for file mode)</param>
    /// <param name="sessionId">Session ID (null for file mode)</param>
    /// <param name="path">Source file path (null for session mode)</param>
    /// <param name="ownsDocument">Whether this context owns and should dispose the document</param>
    private DocumentContext(T document, DocumentSessionManager? sessionManager, string? sessionId, string? path,
        bool ownsDocument)
    {
        Document = document;
        _sessionManager = sessionManager;
        _sessionId = sessionId;
        SourcePath = path;
        _ownsDocument = ownsDocument;
    }

    /// <summary>
    ///     The document instance
    /// </summary>
    public T Document { get; }

    /// <summary>
    ///     Whether this context is from a session
    /// </summary>
    public bool IsSession => _sessionId != null;

    /// <summary>
    ///     The source file path (null for session mode)
    /// </summary>
    public string? SourcePath { get; }

    /// <summary>
    ///     Disposes the document context and releases resources.
    ///     Only disposes the document if it was loaded from file (not from session).
    /// </summary>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_ownsDocument && Document is IDisposable disposable) disposable.Dispose();
    }

    /// <summary>
    ///     Creates a document context from either a session ID or file path
    /// </summary>
    /// <param name="sessionManager">Session manager (can be null if sessions not enabled)</param>
    /// <param name="sessionId">Session ID (if using session mode)</param>
    /// <param name="path">File path (if using file mode)</param>
    /// <returns>Document context</returns>
    public static DocumentContext<T> Create(DocumentSessionManager? sessionManager, string? sessionId, string? path)
    {
        if (!string.IsNullOrEmpty(sessionId))
        {
            if (sessionManager == null)
                throw new InvalidOperationException(
                    "Session management is not enabled. Use --enable-sessions flag or provide a file path.");

            var doc = sessionManager.GetDocument<T>(sessionId);
            return new DocumentContext<T>(doc, sessionManager, sessionId, null, false);
        }

        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Either sessionId or path must be provided");

        SecurityHelper.ValidateFilePath(path, "path", true);

        var document = LoadDocument(path);
        return new DocumentContext<T>(document, null, null, path, true);
    }

    /// <summary>
    ///     Marks the document as modified (for session mode)
    /// </summary>
    public void MarkDirty()
    {
        if (_sessionId != null && _sessionManager != null) _sessionManager.MarkDirty(_sessionId);
    }

    /// <summary>
    ///     Saves the document (only for file mode, session mode just marks dirty)
    /// </summary>
    /// <param name="outputPath">Output path (optional, defaults to source path)</param>
    /// <exception cref="InvalidOperationException">Thrown when no output path is available</exception>
    public void Save(string? outputPath = null)
    {
        if (_sessionId != null)
        {
            MarkDirty();
            return;
        }

        var savePath = outputPath ?? SourcePath ?? throw new InvalidOperationException("No output path available");
        SaveDocument(Document, savePath);
    }

    /// <summary>
    ///     Gets the output message based on context type
    /// </summary>
    /// <param name="outputPath">Output path (optional)</param>
    /// <returns>Message describing where changes were applied or saved</returns>
    public string GetOutputMessage(string? outputPath = null)
    {
        if (_sessionId != null)
            return $"Changes applied to session {_sessionId}. Use document_session(operation='save') to save to disk.";

        return $"Output: {outputPath ?? SourcePath}";
    }

    /// <summary>
    ///     Loads a document from the specified file path based on type T.
    /// </summary>
    /// <param name="path">The file path to load the document from.</param>
    /// <returns>The loaded document instance.</returns>
    /// <exception cref="NotSupportedException">Thrown when the document type is not supported.</exception>
    private static T LoadDocument(string path)
    {
        var type = typeof(T);

        if (type == typeof(Document))
            return (T)(object)new Document(path);

        if (type == typeof(Workbook))
            return (T)(object)new Workbook(path);

        if (type == typeof(Presentation))
            return (T)(object)new Presentation(path);

        if (type == typeof(Aspose.Pdf.Document))
            return (T)(object)new Aspose.Pdf.Document(path);

        throw new NotSupportedException($"Document type {type.Name} is not supported");
    }

    /// <summary>
    ///     Saves the document to the specified file path.
    /// </summary>
    /// <param name="document">The document to save.</param>
    /// <param name="path">The file path to save to.</param>
    /// <exception cref="NotSupportedException">Thrown when the document type is not supported.</exception>
    private static void SaveDocument(T document, string path)
    {
        switch (document)
        {
            case Document wordDoc:
                wordDoc.Save(path);
                break;
            case Workbook workbook:
                workbook.Save(path);
                break;
            case Presentation presentation:
                presentation.Save(path, SaveFormat.Pptx);
                break;
            case Aspose.Pdf.Document pdfDoc:
                pdfDoc.Save(path);
                break;
            default:
                throw new NotSupportedException($"Document type {typeof(T).Name} is not supported");
        }
    }
}