using System.Collections.Concurrent;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using AsposeMcpServer.Helpers;
using SaveFormat = Aspose.Slides.Export.SaveFormat;

namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Manages document sessions for in-memory document editing
/// </summary>
public class DocumentSessionManager : IDisposable
{
    /// <summary>
    ///     Timer for periodic auto-save of dirty sessions
    /// </summary>
    private readonly Timer? _autoSaveTimer;

    /// <summary>
    ///     Timer for periodic cleanup of idle sessions
    /// </summary>
    private readonly Timer? _cleanupTimer;

    /// <summary>
    ///     Logger for session management operations
    /// </summary>
    private readonly ILogger<DocumentSessionManager>? _logger;

    /// <summary>
    ///     Thread-safe dictionary of active sessions grouped by owner key
    ///     Key: Owner storage key, Value: Dictionary of SessionId -> Session
    /// </summary>
    private readonly ConcurrentDictionary<string, ConcurrentDictionary<string, DocumentSession>> _sessionsByOwner =
        new();

    /// <summary>
    ///     Tracks whether this manager has been disposed (0 = not disposed, 1 = disposed)
    /// </summary>
    private int _disposed;

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

        if (config.AutoSaveIntervalMinutes > 0)
            _autoSaveTimer = new Timer(
                AutoSaveDirtySessions,
                null,
                TimeSpan.FromMinutes(config.AutoSaveIntervalMinutes),
                TimeSpan.FromMinutes(config.AutoSaveIntervalMinutes));
    }

    /// <summary>
    ///     Gets the session configuration.
    /// </summary>
    public SessionConfig Config { get; }

    /// <summary>
    ///     Disposes the session manager and all active sessions.
    ///     Thread-safe: uses Interlocked to prevent double-dispose.
    /// </summary>
    public void Dispose()
    {
        // Atomically set _disposed to 1, return previous value
        // If previous value was already 1, another thread already disposed
        if (Interlocked.Exchange(ref _disposed, 1) == 1)
            return;

        _cleanupTimer?.Dispose();
        _autoSaveTimer?.Dispose();

        foreach (var ownerSessions in _sessionsByOwner.Values)
        foreach (var session in ownerSessions.Values)
            session.Dispose();

        _sessionsByOwner.Clear();
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
        return OpenDocument(path, SessionIdentity.GetAnonymous(), mode);
    }

    /// <summary>
    ///     Opens a document and creates a session with owner identity
    /// </summary>
    /// <param name="path">File path to open</param>
    /// <param name="owner">Session owner identity</param>
    /// <param name="mode">Access mode (readonly, readwrite)</param>
    /// <returns>Session ID for the opened document</returns>
    /// <exception cref="ArgumentException">Thrown when mode is invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when maximum session limit reached or file too large</exception>
    /// <exception cref="FileNotFoundException">Thrown when file not found</exception>
    public string OpenDocument(string path, SessionIdentity owner, string mode = "readwrite")
    {
        var normalizedMode = mode.ToLowerInvariant();
        if (normalizedMode != "readonly" && normalizedMode != "readwrite")
            throw new ArgumentException($"Invalid mode: '{mode}'. Must be 'readonly' or 'readwrite'.", nameof(mode));

        var ownerKey = owner.GetStorageKey(Config.IsolationMode);
        var ownerSessions =
            _sessionsByOwner.GetOrAdd(ownerKey, _ => new ConcurrentDictionary<string, DocumentSession>());

        // Early check for session limit (not atomic, but avoids unnecessary work)
        if (ownerSessions.Count >= Config.MaxSessions)
            throw new InvalidOperationException($"Maximum session limit ({Config.MaxSessions}) reached for this user");

        var fileInfo = new FileInfo(path);
        if (!fileInfo.Exists) throw new FileNotFoundException($"File not found: {path}");

        var fileSizeMb = fileInfo.Length / (1024.0 * 1024.0);
        if (fileSizeMb > Config.MaxFileSizeMb)
            throw new InvalidOperationException(
                $"File size ({fileSizeMb:F2} MB) exceeds maximum ({Config.MaxFileSizeMb} MB)");

        var type = GetDocumentType(path);
        var document = LoadDocument(path, type);
        var sessionId = GenerateSessionId();

        var session = new DocumentSession(sessionId, path, type, document, normalizedMode)
        {
            EstimatedMemoryBytes = fileInfo.Length * 2,
            Owner = owner
        };

        if (!ownerSessions.TryAdd(sessionId, session))
        {
            session.Dispose();
            throw new InvalidOperationException("Failed to create session");
        }

        // Atomic check after add: if we exceeded the limit due to race condition, rollback
        if (ownerSessions.Count > Config.MaxSessions)
        {
            if (ownerSessions.TryRemove(sessionId, out _)) session.Dispose();
            throw new InvalidOperationException($"Maximum session limit ({Config.MaxSessions}) reached for this user");
        }

        _logger?.LogInformation("Opened session {SessionId} for {Path} ({Type}) by {Owner}", sessionId, path, type,
            owner);
        return sessionId;
    }

    /// <summary>
    ///     Gets a document from an existing session (no authorization check)
    /// </summary>
    /// <typeparam name="T">Target document type</typeparam>
    /// <param name="sessionId">Session ID to get document from</param>
    /// <returns>Document cast to type T</returns>
    /// <exception cref="KeyNotFoundException">Thrown when session not found</exception>
    public T GetDocument<T>(string sessionId) where T : class
    {
        return GetDocument<T>(sessionId, SessionIdentity.GetAnonymous());
    }

    /// <summary>
    ///     Gets a document from an existing session with authorization check
    /// </summary>
    /// <typeparam name="T">Target document type</typeparam>
    /// <param name="sessionId">Session ID to get document from</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    /// <returns>Document cast to type T</returns>
    /// <exception cref="KeyNotFoundException">Thrown when session not found or access denied</exception>
    public T GetDocument<T>(string sessionId, SessionIdentity requestor) where T : class
    {
        var session = FindSessionWithAuth(sessionId, requestor);
        if (session == null)
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        return session.GetDocument<T>();
    }

    /// <summary>
    ///     Gets a session by ID (no authorization check)
    /// </summary>
    /// <param name="sessionId">Session ID to retrieve</param>
    /// <returns>The document session</returns>
    /// <exception cref="KeyNotFoundException">Thrown when session not found</exception>
    public DocumentSession GetSession(string sessionId)
    {
        return GetSession(sessionId, SessionIdentity.GetAnonymous());
    }

    /// <summary>
    ///     Gets a session by ID with authorization check
    /// </summary>
    /// <param name="sessionId">Session ID to retrieve</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    /// <returns>The document session</returns>
    /// <exception cref="KeyNotFoundException">Thrown when session not found or access denied</exception>
    public DocumentSession GetSession(string sessionId, SessionIdentity requestor)
    {
        var session = FindSessionWithAuth(sessionId, requestor);
        if (session == null)
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        return session;
    }

    /// <summary>
    ///     Tries to get a session by ID with authorization check
    /// </summary>
    /// <param name="sessionId">Session ID to retrieve</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    /// <returns>The document session or null if not found/access denied</returns>
    public DocumentSession? TryGetSession(string sessionId, SessionIdentity requestor)
    {
        return FindSessionWithAuth(sessionId, requestor);
    }

    /// <summary>
    ///     Finds a session by ID and checks authorization
    /// </summary>
    /// <param name="sessionId">Session ID to find</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    /// <returns>Session if found, authorized, and not disposed; null otherwise</returns>
    private DocumentSession? FindSessionWithAuth(string sessionId, SessionIdentity requestor)
    {
        foreach (var ownerSessions in _sessionsByOwner.Values)
            if (ownerSessions.TryGetValue(sessionId, out var session))
            {
                // Check if session is disposed (safety check for race conditions)
                if (session.IsDisposed)
                {
                    _logger?.LogWarning("Attempted to access disposed session {SessionId}", sessionId);
                    // Try to clean up the orphaned entry
                    ownerSessions.TryRemove(sessionId, out _);
                    return null;
                }

                if (!requestor.CanAccess(session.Owner, Config.IsolationMode))
                {
                    _logger?.LogWarning(
                        "Access denied: {Requestor} attempted to access session {SessionId} owned by {Owner}",
                        requestor, sessionId, session.Owner);
                    return null;
                }

                return session;
            }

        return null;
    }

    /// <summary>
    ///     Marks a session as having unsaved changes (no authorization check)
    /// </summary>
    /// <param name="sessionId">Session ID to mark as dirty</param>
    public void MarkDirty(string sessionId)
    {
        MarkDirty(sessionId, SessionIdentity.GetAnonymous());
    }

    /// <summary>
    ///     Marks a session as having unsaved changes with authorization check
    /// </summary>
    /// <param name="sessionId">Session ID to mark as dirty</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    public void MarkDirty(string sessionId, SessionIdentity requestor)
    {
        var session = FindSessionWithAuth(sessionId, requestor);
        if (session != null) session.IsDirty = true;
    }

    /// <summary>
    ///     Saves the document in a session (no authorization check)
    /// </summary>
    /// <param name="sessionId">Session ID to save</param>
    /// <param name="outputPath">Optional output path (defaults to original path)</param>
    /// <exception cref="KeyNotFoundException">Thrown when session not found</exception>
    /// <exception cref="InvalidOperationException">Thrown when trying to save a readonly session</exception>
    public void SaveDocument(string sessionId, string? outputPath = null)
    {
        SaveDocument(sessionId, SessionIdentity.GetAnonymous(), outputPath);
    }

    /// <summary>
    ///     Saves the document in a session with authorization check
    /// </summary>
    /// <param name="sessionId">Session ID to save</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    /// <param name="outputPath">Optional output path (defaults to original path)</param>
    /// <exception cref="KeyNotFoundException">Thrown when session not found or access denied</exception>
    /// <exception cref="InvalidOperationException">Thrown when trying to save a readonly session</exception>
    public void SaveDocument(string sessionId, SessionIdentity requestor, string? outputPath = null)
    {
        var session = FindSessionWithAuth(sessionId, requestor);
        if (session == null)
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        if (session.Mode == "readonly") throw new InvalidOperationException("Cannot save a readonly session");

        var savePath = outputPath ?? session.Path;
        SaveDocumentToFile(session.Document, session.Type, savePath);
        session.IsDirty = false;

        _logger?.LogInformation("Saved session {SessionId} to {Path}", sessionId, savePath);
    }

    /// <summary>
    ///     Closes a session (no authorization check)
    /// </summary>
    /// <param name="sessionId">Session ID to close</param>
    /// <param name="discard">If true, discard unsaved changes; otherwise auto-save</param>
    /// <exception cref="KeyNotFoundException">Thrown when session not found</exception>
    public void CloseDocument(string sessionId, bool discard = false)
    {
        CloseDocument(sessionId, SessionIdentity.GetAnonymous(), discard);
    }

    /// <summary>
    ///     Closes a session with authorization check
    /// </summary>
    /// <param name="sessionId">Session ID to close</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    /// <param name="discard">If true, discard unsaved changes; otherwise auto-save</param>
    /// <exception cref="KeyNotFoundException">Thrown when session not found or access denied</exception>
    public void CloseDocument(string sessionId, SessionIdentity requestor, bool discard = false)
    {
        DocumentSession? session = null;
        ConcurrentDictionary<string, DocumentSession>? ownerDict = null;
        string? ownerKey = null;

        foreach (var kvp in _sessionsByOwner)
            if (kvp.Value.TryGetValue(sessionId, out var foundSession))
            {
                if (!requestor.CanAccess(foundSession.Owner, Config.IsolationMode))
                {
                    _logger?.LogWarning(
                        "Access denied: {Requestor} attempted to close session {SessionId} owned by {Owner}",
                        requestor, sessionId, foundSession.Owner);
                    throw new KeyNotFoundException($"Session not found: {sessionId}");
                }

                if (kvp.Value.TryRemove(sessionId, out session))
                {
                    ownerDict = kvp.Value;
                    ownerKey = kvp.Key;
                }

                break;
            }

        if (session == null)
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        // Safely clean up empty owner dictionary using atomic operation
        if (ownerDict != null && ownerKey != null && ownerDict.IsEmpty)
            ((ICollection<KeyValuePair<string, ConcurrentDictionary<string, DocumentSession>>>)_sessionsByOwner)
                .Remove(new KeyValuePair<string, ConcurrentDictionary<string, DocumentSession>>(ownerKey, ownerDict));

        try
        {
            if (!discard && session.IsDirty)
                SaveDocumentToFile(session.Document, session.Type, session.Path);
        }
        finally
        {
            // Always dispose session to prevent resource leaks, even if save fails
            session.Dispose();
        }

        _logger?.LogInformation("Closed session {SessionId} (discard={Discard})", sessionId, discard);
    }

    /// <summary>
    ///     Lists all active sessions (no authorization check - returns all)
    /// </summary>
    /// <returns>Enumerable of session information</returns>
    public IEnumerable<SessionInfo> ListSessions()
    {
        return ListSessions(SessionIdentity.GetAnonymous());
    }

    /// <summary>
    ///     Lists active sessions visible to the requestor
    /// </summary>
    /// <param name="requestor">Requestor identity for filtering</param>
    /// <returns>Enumerable of session information</returns>
    public IEnumerable<SessionInfo> ListSessions(SessionIdentity requestor)
    {
        IEnumerable<DocumentSession> sessions;

        if (Config.IsolationMode == SessionIsolationMode.None)
        {
            sessions = _sessionsByOwner.Values.SelectMany(s => s.Values);
        }
        else
        {
            var ownerKey = requestor.GetStorageKey(Config.IsolationMode);
            sessions = _sessionsByOwner.TryGetValue(ownerKey, out var ownerSessions)
                ? ownerSessions.Values
                : [];
        }

        return sessions.Select(s => new SessionInfo
        {
            SessionId = s.SessionId,
            DocumentType = s.Type.ToString().ToLowerInvariant(),
            Path = s.Path,
            Mode = s.Mode,
            IsDirty = s.IsDirty,
            OpenedAt = s.OpenedAt,
            LastAccessedAt = s.LastAccessedAt,
            EstimatedMemoryMb = s.EstimatedMemoryBytes / (1024.0 * 1024.0)
        });
    }

    /// <summary>
    ///     Gets the status of a specific session (no authorization check)
    /// </summary>
    /// <param name="sessionId">Session ID to get status for</param>
    /// <returns>Session information or null if not found</returns>
    public SessionInfo? GetSessionStatus(string sessionId)
    {
        return GetSessionStatus(sessionId, SessionIdentity.GetAnonymous());
    }

    /// <summary>
    ///     Gets the status of a specific session with authorization check
    /// </summary>
    /// <param name="sessionId">Session ID to get status for</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    /// <returns>Session information or null if not found/access denied</returns>
    public SessionInfo? GetSessionStatus(string sessionId, SessionIdentity requestor)
    {
        var session = FindSessionWithAuth(sessionId, requestor);
        if (session == null) return null;

        return new SessionInfo
        {
            SessionId = session.SessionId,
            DocumentType = session.Type.ToString().ToLowerInvariant(),
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
        return _sessionsByOwner.Values
            .SelectMany(s => s.Values)
            .Sum(s => s.EstimatedMemoryBytes) / (1024.0 * 1024.0);
    }

    /// <summary>
    ///     Handles server shutdown - saves or discards sessions based on config
    /// </summary>
    public void OnServerShutdown()
    {
        var allSessions = _sessionsByOwner.Values.SelectMany(s => s.Values).ToList();
        _logger?.LogInformation("Server shutdown - handling {Count} open sessions", allSessions.Count);

        foreach (var session in allSessions)
            try
            {
                HandleDisconnect(session);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error handling session {SessionId} on shutdown", session.SessionId);
            }

        _sessionsByOwner.Clear();
    }

    /// <summary>
    ///     Handles client disconnect - saves or discards sessions based on config
    /// </summary>
    /// <param name="clientId">Client identifier to disconnect</param>
    public void OnClientDisconnect(string? clientId)
    {
        if (string.IsNullOrEmpty(clientId)) return;

        var clientSessions = _sessionsByOwner.Values
            .SelectMany(s => s.Values)
            .Where(s => s.ClientId == clientId)
            .ToList();

        _logger?.LogInformation("Client {ClientId} disconnected - handling {Count} sessions", clientId,
            clientSessions.Count);

        foreach (var session in clientSessions)
            try
            {
                // First remove from storage, then handle disconnect (dispose)
                // This prevents race conditions where disposed sessions are still accessible
                RemoveSessionById(session.SessionId);
                HandleDisconnect(session);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error handling session {SessionId} on disconnect", session.SessionId);
            }
    }

    /// <summary>
    ///     Removes a session by ID from storage (no authorization check)
    /// </summary>
    /// <param name="sessionId">Session ID to remove</param>
    private void RemoveSessionById(string sessionId)
    {
        foreach (var kvp in _sessionsByOwner)
            if (kvp.Value.TryRemove(sessionId, out _))
            {
                // Safely clean up empty owner dictionary using atomic operation
                // Only remove if the dictionary is still empty (handles race condition)
                if (kvp.Value.IsEmpty)
                    // Use TryUpdate pattern: only remove if value matches (empty dictionary)
                    ((ICollection<KeyValuePair<string, ConcurrentDictionary<string, DocumentSession>>>)_sessionsByOwner)
                        .Remove(new KeyValuePair<string, ConcurrentDictionary<string, DocumentSession>>(kvp.Key,
                            kvp.Value));
                return;
            }
    }

    /// <summary>
    ///     Handles disconnect behavior for a session based on configuration.
    ///     Always disposes the session, even if save fails or session is not dirty.
    /// </summary>
    /// <param name="session">Session to handle disconnect for</param>
    private void HandleDisconnect(DocumentSession session)
    {
        try
        {
            if (!session.IsDirty)
            {
                _logger?.LogDebug("Session {SessionId} has no unsaved changes", session.SessionId);
                return;
            }

            switch (Config.OnDisconnect)
            {
                case DisconnectBehavior.AutoSave:
                    SaveDocumentToFile(session.Document, session.Type, session.Path);
                    DeleteSessionTempFiles(session.SessionId);
                    _logger?.LogInformation("Auto-saved session {SessionId} and cleaned up temp files",
                        session.SessionId);
                    break;

                case DisconnectBehavior.SaveToTemp:
                    var tempPath = GetTempPath(session);
                    SaveDocumentToFile(session.Document, session.Type, tempPath);
                    SaveSessionMetadata(session, tempPath);
                    _logger?.LogInformation("Saved session {SessionId} to temp: {TempPath}", session.SessionId,
                        tempPath);
                    break;

                case DisconnectBehavior.Discard:
                    DeleteSessionTempFiles(session.SessionId);
                    _logger?.LogInformation("Discarded changes for session {SessionId} and cleaned up temp files",
                        session.SessionId);
                    break;

                case DisconnectBehavior.PromptOnReconnect:
                    var promptTempPath = GetTempPath(session);
                    SaveDocumentToFile(session.Document, session.Type, promptTempPath);
                    SaveSessionMetadata(session, promptTempPath, true);
                    _logger?.LogInformation("Saved session {SessionId} for prompt on reconnect", session.SessionId);
                    break;
            }
        }
        finally
        {
            // Always dispose session to prevent resource leaks
            session.Dispose();
        }
    }

    /// <summary>
    ///     Timer callback to cleanup idle sessions
    /// </summary>
    /// <param name="state">Timer state (not used)</param>
    private void CleanupIdleSessions(object? state)
    {
        var timeout = TimeSpan.FromMinutes(Config.IdleTimeoutMinutes);
        var now = DateTime.UtcNow;

        var allSessions = _sessionsByOwner.Values
            .SelectMany(s => s.Values)
            .ToList();

        foreach (var session in allSessions)
            if (now - session.LastAccessedAt > timeout)
            {
                _logger?.LogInformation("Session {SessionId} timed out after {Minutes} minutes of inactivity",
                    session.SessionId, Config.IdleTimeoutMinutes);

                try
                {
                    // First remove from storage, then handle disconnect (dispose)
                    // This prevents race conditions where disposed sessions are still accessible
                    RemoveSessionById(session.SessionId);
                    HandleDisconnect(session);
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "Error cleaning up idle session {SessionId}", session.SessionId);
                }
            }
    }

    /// <summary>
    ///     Timer callback to auto-save dirty sessions to temp files.
    ///     This helps prevent data loss in case of unexpected termination (e.g., kill -9).
    ///     Unlike HandleDisconnect, this does NOT dispose or remove sessions - they remain active.
    /// </summary>
    /// <param name="state">Timer state (not used)</param>
    private void AutoSaveDirtySessions(object? state)
    {
        var dirtySessions = _sessionsByOwner.Values
            .SelectMany(s => s.Values)
            .Where(s => s is { IsDirty: true, IsDisposed: false })
            .ToList();

        if (dirtySessions.Count == 0)
            return;

        _logger?.LogDebug("Auto-saving {Count} dirty sessions", dirtySessions.Count);

        foreach (var session in dirtySessions)
            try
            {
                var tempPath = GetTempPath(session);
                SaveDocumentToFile(session.Document, session.Type, tempPath);
                SaveSessionMetadata(session, tempPath);
                _logger?.LogInformation("Auto-saved dirty session {SessionId} to temp: {TempPath}",
                    session.SessionId, tempPath);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error auto-saving session {SessionId}", session.SessionId);
            }
    }

    /// <summary>
    ///     Generates a unique session ID
    /// </summary>
    /// <returns>Generated session ID in format sess_XXXXXXXXXXXXXXXXXXXXXXXX (24 chars total)</returns>
    private static string GenerateSessionId()
    {
        // Use 19 hex chars from GUID for better uniqueness (76 bits of randomness)
        // Format: sess_ (5) + 19 hex chars = 24 chars total
        return $"sess_{Guid.NewGuid():N}"[..24];
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
    ///     Deletes all temporary files associated with a session.
    ///     This is called when AutoSave or Discard behavior is used to prevent
    ///     stale temp files from appearing in list_temp_files.
    /// </summary>
    /// <param name="sessionId">Session ID to delete temp files for</param>
    private void DeleteSessionTempFiles(string sessionId)
    {
        try
        {
            var pattern = $"aspose_session_{sessionId}_*.json";
            var metadataFiles = Directory.GetFiles(Config.TempDirectory, pattern);

            foreach (var metadataPath in metadataFiles)
                try
                {
                    var json = File.ReadAllText(metadataPath);
                    var metadata = JsonSerializer.Deserialize<TempFileMetadata>(json);

                    if (metadata?.TempPath != null && File.Exists(metadata.TempPath))
                        File.Delete(metadata.TempPath);

                    File.Delete(metadataPath);
                    _logger?.LogDebug("Deleted temp file for session {SessionId}: {Path}", sessionId, metadataPath);
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "Failed to delete temp file: {Path}", metadataPath);
                }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to cleanup temp files for session {SessionId}", sessionId);
        }
    }

    /// <summary>
    ///     Saves session metadata for recovery purposes
    /// </summary>
    /// <param name="session">Session to save metadata for</param>
    /// <param name="tempPath">Temporary file path where document was saved</param>
    /// <param name="promptOnReconnect">Whether to prompt user on reconnect</param>
    private static void SaveSessionMetadata(DocumentSession session, string tempPath, bool promptOnReconnect = false)
    {
        var metadata = new
        {
            session.SessionId,
            OriginalPath = session.Path,
            TempPath = tempPath,
            DocumentType = session.Type.ToString(),
            SavedAt = DateTime.UtcNow,
            PromptOnReconnect = promptOnReconnect,
            OwnerGroupId = session.Owner.GroupId,
            OwnerUserId = session.Owner.UserId
        };

        var metadataPath = tempPath + ".meta.json";
        File.WriteAllText(metadataPath, JsonSerializer.Serialize(metadata, JsonDefaults.Indented));
    }
}
