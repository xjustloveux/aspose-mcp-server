using System.Collections.Concurrent;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Manages document sessions for in-memory document editing
/// </summary>
public class DocumentSessionManager : IDisposable
{
    /// <summary>
    ///     Timer for periodic cleanup of idle sessions
    /// </summary>
    private readonly Timer? _cleanupTimer;

    /// <summary>
    ///     Logger for session management operations
    /// </summary>
    private readonly ILogger<DocumentSessionManager>? _logger;

    /// <summary>
    ///     Thread-safe dictionary of active sessions
    /// </summary>
    private readonly ConcurrentDictionary<string, DocumentSession> _sessions = new();

    /// <summary>
    ///     Tracks whether this manager has been disposed
    /// </summary>
    private bool _disposed;

    /// <summary>
    ///     Creates a new document session manager
    /// </summary>
    /// <param name="config">Session configuration</param>
    /// <param name="loggerFactory">Logger factory for logging</param>
    public DocumentSessionManager(SessionConfig config, ILoggerFactory? loggerFactory = null)
    {
        Config = config;
        _logger = loggerFactory?.CreateLogger<DocumentSessionManager>();

        if (config.IdleTimeoutMinutes > 0)
            _cleanupTimer = new Timer(
                CleanupIdleSessions,
                null,
                TimeSpan.FromMinutes(1),
                TimeSpan.FromMinutes(1));
    }

    /// <summary>
    ///     Gets the session configuration.
    /// </summary>
    public SessionConfig Config { get; }

    /// <summary>
    ///     Disposes the session manager and all active sessions.
    /// </summary>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        _cleanupTimer?.Dispose();

        foreach (var session in _sessions.Values) session.Dispose();

        _sessions.Clear();
    }

    /// <summary>
    ///     Opens a document and creates a session
    /// </summary>
    /// <param name="path">File path to open</param>
    /// <param name="mode">Access mode (readonly, readwrite)</param>
    /// <returns>Session ID for the opened document</returns>
    /// <exception cref="InvalidOperationException">Thrown when maximum session limit reached or file too large</exception>
    /// <exception cref="FileNotFoundException">Thrown when file not found</exception>
    public string OpenDocument(string path, string mode = "readwrite")
    {
        if (_sessions.Count >= Config.MaxSessions)
            throw new InvalidOperationException($"Maximum session limit ({Config.MaxSessions}) reached");

        var fileInfo = new FileInfo(path);
        if (!fileInfo.Exists) throw new FileNotFoundException($"File not found: {path}");

        var fileSizeMb = fileInfo.Length / (1024.0 * 1024.0);
        if (fileSizeMb > Config.MaxFileSizeMb)
            throw new InvalidOperationException(
                $"File size ({fileSizeMb:F2} MB) exceeds maximum ({Config.MaxFileSizeMb} MB)");

        var type = GetDocumentType(path);
        var document = LoadDocument(path, type);
        var sessionId = GenerateSessionId();

        var session = new DocumentSession(sessionId, path, type, document, mode)
        {
            EstimatedMemoryBytes = fileInfo.Length * 2
        };

        if (!_sessions.TryAdd(sessionId, session))
        {
            session.Dispose();
            throw new InvalidOperationException("Failed to create session");
        }

        _logger?.LogInformation("Opened session {SessionId} for {Path} ({Type})", sessionId, path, type);
        return sessionId;
    }

    /// <summary>
    ///     Gets a document from an existing session
    /// </summary>
    /// <typeparam name="T">Target document type</typeparam>
    /// <param name="sessionId">Session ID to get document from</param>
    /// <returns>Document cast to type T</returns>
    /// <exception cref="KeyNotFoundException">Thrown when session not found</exception>
    public T GetDocument<T>(string sessionId) where T : class
    {
        if (!_sessions.TryGetValue(sessionId, out var session))
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        return session.GetDocument<T>();
    }

    /// <summary>
    ///     Gets a session by ID
    /// </summary>
    /// <param name="sessionId">Session ID to retrieve</param>
    /// <returns>The document session</returns>
    /// <exception cref="KeyNotFoundException">Thrown when session not found</exception>
    public DocumentSession GetSession(string sessionId)
    {
        if (!_sessions.TryGetValue(sessionId, out var session))
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        return session;
    }

    /// <summary>
    ///     Marks a session as having unsaved changes
    /// </summary>
    /// <param name="sessionId">Session ID to mark as dirty</param>
    public void MarkDirty(string sessionId)
    {
        if (_sessions.TryGetValue(sessionId, out var session)) session.IsDirty = true;
    }

    /// <summary>
    ///     Saves the document in a session
    /// </summary>
    /// <param name="sessionId">Session ID to save</param>
    /// <param name="outputPath">Optional output path (defaults to original path)</param>
    /// <exception cref="KeyNotFoundException">Thrown when session not found</exception>
    /// <exception cref="InvalidOperationException">Thrown when trying to save a readonly session</exception>
    public void SaveDocument(string sessionId, string? outputPath = null)
    {
        if (!_sessions.TryGetValue(sessionId, out var session))
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        if (session.Mode == "readonly") throw new InvalidOperationException("Cannot save a readonly session");

        var savePath = outputPath ?? session.Path;
        SaveDocumentToFile(session.Document, session.Type, savePath);
        session.IsDirty = false;

        _logger?.LogInformation("Saved session {SessionId} to {Path}", sessionId, savePath);
    }

    /// <summary>
    ///     Closes a session
    /// </summary>
    /// <param name="sessionId">Session ID to close</param>
    /// <param name="discard">If true, discard unsaved changes; otherwise auto-save</param>
    /// <exception cref="KeyNotFoundException">Thrown when session not found</exception>
    public void CloseDocument(string sessionId, bool discard = false)
    {
        if (!_sessions.TryRemove(sessionId, out var session))
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        if (!discard && session.IsDirty) SaveDocumentToFile(session.Document, session.Type, session.Path);

        session.Dispose();
        _logger?.LogInformation("Closed session {SessionId} (discard={Discard})", sessionId, discard);
    }

    /// <summary>
    ///     Lists all active sessions
    /// </summary>
    /// <returns>Enumerable of session information</returns>
    public IEnumerable<SessionInfo> ListSessions()
    {
        return _sessions.Values.Select(s => new SessionInfo
        {
            SessionId = s.SessionId,
            DocumentType = s.Type.ToString().ToLower(),
            Path = s.Path,
            Mode = s.Mode,
            IsDirty = s.IsDirty,
            OpenedAt = s.OpenedAt,
            LastAccessedAt = s.LastAccessedAt,
            EstimatedMemoryMb = s.EstimatedMemoryBytes / (1024.0 * 1024.0)
        });
    }

    /// <summary>
    ///     Gets the status of a specific session
    /// </summary>
    /// <param name="sessionId">Session ID to get status for</param>
    /// <returns>Session information or null if not found</returns>
    public SessionInfo? GetSessionStatus(string sessionId)
    {
        if (!_sessions.TryGetValue(sessionId, out var session)) return null;

        return new SessionInfo
        {
            SessionId = session.SessionId,
            DocumentType = session.Type.ToString().ToLower(),
            Path = session.Path,
            Mode = session.Mode,
            IsDirty = session.IsDirty,
            OpenedAt = session.OpenedAt,
            LastAccessedAt = session.LastAccessedAt,
            EstimatedMemoryMb = session.EstimatedMemoryBytes / (1024.0 * 1024.0)
        };
    }

    /// <summary>
    ///     Gets total memory used by all sessions
    /// </summary>
    /// <returns>Total memory usage in megabytes</returns>
    public double GetTotalMemoryMb()
    {
        return _sessions.Values.Sum(s => s.EstimatedMemoryBytes) / (1024.0 * 1024.0);
    }

    /// <summary>
    ///     Handles server shutdown - saves or discards sessions based on config
    /// </summary>
    public void OnServerShutdown()
    {
        _logger?.LogInformation("Server shutdown - handling {Count} open sessions", _sessions.Count);

        foreach (var session in _sessions.Values)
            try
            {
                HandleDisconnect(session);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error handling session {SessionId} on shutdown", session.SessionId);
            }

        _sessions.Clear();
    }

    /// <summary>
    ///     Handles client disconnect - saves or discards sessions based on config
    /// </summary>
    /// <param name="clientId">Client identifier to disconnect</param>
    public void OnClientDisconnect(string? clientId)
    {
        if (string.IsNullOrEmpty(clientId)) return;

        var clientSessions = _sessions.Values.Where(s => s.ClientId == clientId).ToList();
        _logger?.LogInformation("Client {ClientId} disconnected - handling {Count} sessions", clientId,
            clientSessions.Count);

        foreach (var session in clientSessions)
            try
            {
                HandleDisconnect(session);
                _sessions.TryRemove(session.SessionId, out _);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error handling session {SessionId} on disconnect", session.SessionId);
            }
    }

    /// <summary>
    ///     Handles disconnect behavior for a session based on configuration
    /// </summary>
    /// <param name="session">Session to handle disconnect for</param>
    private void HandleDisconnect(DocumentSession session)
    {
        if (!session.IsDirty) return;

        switch (Config.OnDisconnect)
        {
            case DisconnectBehavior.AutoSave:
                SaveDocumentToFile(session.Document, session.Type, session.Path);
                _logger?.LogInformation("Auto-saved session {SessionId}", session.SessionId);
                break;

            case DisconnectBehavior.SaveToTemp:
                var tempPath = GetTempPath(session);
                SaveDocumentToFile(session.Document, session.Type, tempPath);
                SaveSessionMetadata(session, tempPath);
                _logger?.LogInformation("Saved session {SessionId} to temp: {TempPath}", session.SessionId, tempPath);
                break;

            case DisconnectBehavior.Discard:
                _logger?.LogInformation("Discarded changes for session {SessionId}", session.SessionId);
                break;

            case DisconnectBehavior.PromptOnReconnect:
                var promptTempPath = GetTempPath(session);
                SaveDocumentToFile(session.Document, session.Type, promptTempPath);
                SaveSessionMetadata(session, promptTempPath, true);
                _logger?.LogInformation("Saved session {SessionId} for prompt on reconnect", session.SessionId);
                break;
        }

        session.Dispose();
    }

    /// <summary>
    ///     Timer callback to cleanup idle sessions
    /// </summary>
    /// <param name="state">Timer state (not used)</param>
    private void CleanupIdleSessions(object? state)
    {
        var timeout = TimeSpan.FromMinutes(Config.IdleTimeoutMinutes);
        var now = DateTime.UtcNow;

        foreach (var session in _sessions.Values.ToList())
            if (now - session.LastAccessedAt > timeout)
            {
                _logger?.LogInformation("Session {SessionId} timed out after {Minutes} minutes of inactivity",
                    session.SessionId, Config.IdleTimeoutMinutes);

                try
                {
                    HandleDisconnect(session);
                    _sessions.TryRemove(session.SessionId, out _);
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "Error cleaning up idle session {SessionId}", session.SessionId);
                }
            }
    }

    /// <summary>
    ///     Generates a unique session ID
    /// </summary>
    /// <returns>Generated session ID in format sess_XXXXXXXXXXXX</returns>
    private static string GenerateSessionId()
    {
        return $"sess_{Guid.NewGuid():N}"[..16];
    }

    /// <summary>
    ///     Determines document type from file extension
    /// </summary>
    /// <param name="path">File path to determine type for</param>
    /// <returns>Document type enum value</returns>
    /// <exception cref="NotSupportedException">Thrown when file extension is not supported</exception>
    private static DocumentType GetDocumentType(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".doc" or ".docx" or ".docm" or ".dot" or ".dotx" or ".dotm" or ".rtf" or ".odt" => DocumentType.Word,
            ".xls" or ".xlsx" or ".xlsm" or ".xlsb" or ".csv" or ".ods" => DocumentType.Excel,
            ".ppt" or ".pptx" or ".pptm" or ".pot" or ".potx" or ".potm" or ".odp" => DocumentType.PowerPoint,
            ".pdf" => DocumentType.Pdf,
            _ => throw new NotSupportedException($"Unsupported file extension: {ext}")
        };
    }

    /// <summary>
    ///     Loads a document from file based on type
    /// </summary>
    /// <param name="path">File path to load</param>
    /// <param name="type">Document type to load as</param>
    /// <returns>Loaded Aspose document object</returns>
    /// <exception cref="NotSupportedException">Thrown when document type is not supported</exception>
    private static object LoadDocument(string path, DocumentType type)
    {
        return type switch
        {
            DocumentType.Word => new Document(path),
            DocumentType.Excel => new Workbook(path),
            DocumentType.PowerPoint => new Presentation(path),
            DocumentType.Pdf => new Aspose.Pdf.Document(path),
            _ => throw new NotSupportedException($"Unsupported document type: {type}")
        };
    }

    /// <summary>
    ///     Saves a document to file based on type
    /// </summary>
    /// <param name="document">Aspose document object to save</param>
    /// <param name="type">Document type</param>
    /// <param name="path">File path to save to</param>
    private static void SaveDocumentToFile(object document, DocumentType type, string path)
    {
        switch (type)
        {
            case DocumentType.Word:
                ((Document)document).Save(path);
                break;
            case DocumentType.Excel:
                ((Workbook)document).Save(path);
                break;
            case DocumentType.PowerPoint:
                ((Presentation)document).Save(path, SaveFormat.Pptx);
                break;
            case DocumentType.Pdf:
                ((Aspose.Pdf.Document)document).Save(path);
                break;
        }
    }

    /// <summary>
    ///     Generates a temporary file path for session recovery
    /// </summary>
    /// <param name="session">Session to generate temp path for</param>
    /// <returns>Full path to temporary file</returns>
    private string GetTempPath(DocumentSession session)
    {
        var ext = Path.GetExtension(session.Path);
        var timestamp = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
        return Path.Combine(Config.TempDirectory, $"aspose_session_{session.SessionId}_{timestamp}{ext}");
    }

    /// <summary>
    ///     Saves session metadata for recovery purposes
    /// </summary>
    /// <param name="session">Session to save metadata for</param>
    /// <param name="tempPath">Temporary file path where document was saved</param>
    /// <param name="promptOnReconnect">Whether to prompt user on reconnect</param>
    private void SaveSessionMetadata(DocumentSession session, string tempPath, bool promptOnReconnect = false)
    {
        var metadata = new
        {
            session.SessionId,
            OriginalPath = session.Path,
            TempPath = tempPath,
            DocumentType = session.Type.ToString(),
            SavedAt = DateTime.UtcNow,
            PromptOnReconnect = promptOnReconnect
        };

        var metadataPath = tempPath + ".meta.json";
        File.WriteAllText(metadataPath,
            JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true }));
    }
}

/// <summary>
///     Session information for API responses
/// </summary>
public class SessionInfo
{
    /// <summary>
    ///     Unique session identifier
    /// </summary>
    public string SessionId { get; set; } = "";

    /// <summary>
    ///     Document type (word, excel, powerpoint, pdf)
    /// </summary>
    public string DocumentType { get; set; } = "";

    /// <summary>
    ///     Original file path
    /// </summary>
    public string Path { get; set; } = "";

    /// <summary>
    ///     Access mode (readonly, readwrite)
    /// </summary>
    public string Mode { get; set; } = "";

    /// <summary>
    ///     Whether the document has unsaved changes
    /// </summary>
    public bool IsDirty { get; set; }

    /// <summary>
    ///     When the session was opened
    /// </summary>
    public DateTime OpenedAt { get; set; }

    /// <summary>
    ///     Last access time
    /// </summary>
    public DateTime LastAccessedAt { get; set; }

    /// <summary>
    ///     Estimated memory usage in MB
    /// </summary>
    public double EstimatedMemoryMb { get; set; }
}