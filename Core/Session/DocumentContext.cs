using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Core.Helpers;
using LoadOptions = Aspose.Words.Loading.LoadOptions;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Provides unified document access whether from file path or session
/// </summary>
public class DocumentContext<T> : IDisposable where T : class
{
    /// <summary>
    ///     The identity of the requestor for session isolation
    /// </summary>
    private readonly SessionIdentity _identity;

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
    /// <param name="identity">The requestor identity for session isolation</param>
    private DocumentContext(T document, DocumentSessionManager? sessionManager, string? sessionId, string? path,
        bool ownsDocument, SessionIdentity identity)
    {
        Document = document;
        _sessionManager = sessionManager;
        _sessionId = sessionId;
        SourcePath = path;
        _ownsDocument = ownsDocument;
        _identity = identity;
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
    /// <param name="identityAccessor">
    ///     Identity accessor for session isolation (required for session mode with isolation
    ///     enabled)
    /// </param>
    /// <param name="password">Optional password for opening encrypted documents</param>
    /// <returns>Document context</returns>
    public static DocumentContext<T> Create(
        DocumentSessionManager? sessionManager,
        string? sessionId,
        string? path,
        ISessionIdentityAccessor? identityAccessor = null,
        string? password = null)
    {
        var identity = identityAccessor?.GetCurrentIdentity() ?? SessionIdentity.GetAnonymous();

        if (!string.IsNullOrEmpty(sessionId))
        {
            if (sessionManager == null)
                throw new InvalidOperationException(
                    "Session management is not enabled. Use --enable-sessions flag or provide a file path.");

            var doc = sessionManager.GetDocument<T>(sessionId, identity);
            return new DocumentContext<T>(doc, sessionManager, sessionId, null, false, identity);
        }

        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("Either sessionId or path must be provided");

        SecurityHelper.ValidateFilePath(path, "path", true);

        var document = LoadDocument(path, password);
        return new DocumentContext<T>(document, null, null, path, true, identity);
    }

    /// <summary>
    ///     Marks the document as modified (for session mode)
    /// </summary>
    public void MarkDirty()
    {
        if (_sessionId != null && _sessionManager != null) _sessionManager.MarkDirty(_sessionId, _identity);
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
    /// <param name="password">Optional password for encrypted documents.</param>
    /// <returns>The loaded document instance.</returns>
    /// <exception cref="NotSupportedException">Thrown when the document type is not supported.</exception>
    private static T LoadDocument(string path, string? password)
    {
        var type = typeof(T);

        if (type == typeof(Document))
            return (T)(object)LoadWordDocument(path, password);

        if (type == typeof(Workbook))
            return (T)(object)LoadExcelWorkbook(path, password);

        if (type == typeof(Presentation))
            return (T)(object)LoadPowerPointPresentation(path, password);

        if (type == typeof(Aspose.Pdf.Document))
            return (T)(object)LoadPdfDocument(path, password);

        throw new NotSupportedException($"Document type {type.Name} is not supported");
    }

    /// <summary>
    ///     Loads a Word document with optional password support.
    /// </summary>
    private static Document LoadWordDocument(string path, string? password)
    {
        if (!string.IsNullOrEmpty(password))
            try
            {
                var loadOptions = new LoadOptions { Password = password };
                return new Document(path, loadOptions);
            }
            catch (IncorrectPasswordException)
            {
                // Password didn't work for opening, try without password
                // (file might not be encrypted, password is for other operations like protection)
            }

        return new Document(path);
    }

    /// <summary>
    ///     Loads an Excel workbook with optional password support.
    /// </summary>
    private static Workbook LoadExcelWorkbook(string path, string? password)
    {
        if (!string.IsNullOrEmpty(password))
            try
            {
                var loadOptions = new Aspose.Cells.LoadOptions { Password = password };
                return new Workbook(path, loadOptions);
            }
            catch (CellsException)
            {
                // Password didn't work for opening, try without password
            }

        return new Workbook(path);
    }

    /// <summary>
    ///     Loads a PowerPoint presentation with optional password support.
    /// </summary>
    private static Presentation LoadPowerPointPresentation(string path, string? password)
    {
        if (!string.IsNullOrEmpty(password))
            try
            {
                var loadOptions = new Aspose.Slides.LoadOptions { Password = password };
                return new Presentation(path, loadOptions);
            }
            catch (InvalidPasswordException)
            {
                // Password didn't work for opening, try without password
            }

        return new Presentation(path);
    }

    /// <summary>
    ///     Loads a PDF document with optional password support.
    /// </summary>
    private static Aspose.Pdf.Document LoadPdfDocument(string path, string? password)
    {
        if (!string.IsNullOrEmpty(password))
            try
            {
                return new Aspose.Pdf.Document(path, password);
            }
            catch (Aspose.Pdf.InvalidPasswordException)
            {
                // Password didn't work for opening, try without password
            }

        return new Aspose.Pdf.Document(path);
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