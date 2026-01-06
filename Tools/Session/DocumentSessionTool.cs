using System.ComponentModel;
using System.Text.Json;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Session;

/// <summary>
///     Tool for managing document sessions (open, save, close, list, status, temp file recovery)
/// </summary>
[McpServerToolType]
public class DocumentSessionTool
{
    /// <summary>
    ///     The session identity accessor for getting current user identity.
    /// </summary>
    private readonly ISessionIdentityAccessor _identityAccessor;

    /// <summary>
    ///     The document session manager for managing document lifecycle.
    /// </summary>
    private readonly DocumentSessionManager _sessionManager;

    /// <summary>
    ///     The temp file manager for managing temporary files and recovery.
    /// </summary>
    private readonly TempFileManager _tempFileManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="DocumentSessionTool" /> class.
    /// </summary>
    /// <param name="sessionManager">The document session manager instance.</param>
    /// <param name="tempFileManager">The temp file manager instance.</param>
    /// <param name="identityAccessor">The session identity accessor instance.</param>
    public DocumentSessionTool(
        DocumentSessionManager sessionManager,
        TempFileManager tempFileManager,
        ISessionIdentityAccessor identityAccessor)
    {
        _sessionManager = sessionManager;
        _tempFileManager = tempFileManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Manages document sessions for in-memory editing
    /// </summary>
    /// <param name="operation">Operation to perform</param>
    /// <param name="path">File path (for 'open' operation)</param>
    /// <param name="sessionId">Session ID (for session-specific operations)</param>
    /// <param name="outputPath">Output path (for 'save' or 'recover' operations)</param>
    /// <param name="mode">Access mode (for 'open' operation)</param>
    /// <param name="discard">Discard changes (for 'close' operation)</param>
    /// <param name="deleteAfterRecover">Delete temp file after recovery (for 'recover' operation)</param>
    /// <returns>Operation result as JSON string</returns>
    [McpServerTool(Name = "document_session")]
    [Description(@"Manage document sessions for in-memory editing. Supports Word, Excel, PowerPoint, and PDF files.

Session Operations:
- 'open': Open a document and create a session. Returns sessionId for subsequent operations.
- 'save': Save document changes to file (original path or specified outputPath).
- 'close': Close session and optionally discard changes.
- 'list': List all active sessions.
- 'status': Get status of a specific session.

Temp File Operations (for disconnected sessions):
- 'list_temp': List recoverable temp files from previous sessions.
- 'recover': Recover a temp file to original or specified path.
- 'delete_temp': Delete a specific temp file.
- 'cleanup': Clean up expired temp files.
- 'temp_stats': Get temp file statistics.

Usage examples:
- Open document: document_session(operation='open', path='doc.docx', mode='readwrite')
- Save document: document_session(operation='save', sessionId='sess_abc123')
- Save as: document_session(operation='save', sessionId='sess_abc123', outputPath='new.docx')
- Close (save changes): document_session(operation='close', sessionId='sess_abc123')
- Close (discard): document_session(operation='close', sessionId='sess_abc123', discard=true)
- List sessions: document_session(operation='list')
- Get status: document_session(operation='status', sessionId='sess_abc123')
- List temp files: document_session(operation='list_temp')
- Recover temp file: document_session(operation='recover', sessionId='sess_abc123')
- Recover to path: document_session(operation='recover', sessionId='sess_abc123', outputPath='recovered.docx')

After opening a document, use the returned sessionId with other tools to edit the document in memory.
Changes are only written to disk when you call 'save' or 'close' (without discard=true).")]
    public string Execute(
        [Description(@"Operation to perform:
Session operations:
- 'open': Open document and create session (required: path)
- 'save': Save session to file (required: sessionId)
- 'close': Close session (required: sessionId)
- 'list': List all active sessions
- 'status': Get session status (required: sessionId)
Temp file operations:
- 'list_temp': List recoverable temp files
- 'recover': Recover temp file (required: sessionId)
- 'delete_temp': Delete temp file (required: sessionId)
- 'cleanup': Clean up expired temp files
- 'temp_stats': Get temp file statistics")]
        string operation,
        [Description("File path to open (required for 'open' operation)")]
        string? path = null,
        [Description("Session ID (required for 'save', 'close', 'status', 'recover', 'delete_temp' operations)")]
        string? sessionId = null,
        [Description("Output path for save/recover (optional, defaults to original path)")]
        string? outputPath = null,
        [Description("Access mode: 'readonly' or 'readwrite' (for 'open', default: 'readwrite')")]
        string mode = "readwrite",
        [Description("Discard changes when closing (for 'close', default: false)")]
        bool discard = false,
        [Description("Delete temp file after recovery (for 'recover', default: true)")]
        bool deleteAfterRecover = true)
    {
        return operation.ToLower() switch
        {
            // Session operations
            "open" => OpenDocument(path!, mode),
            "save" => SaveDocument(sessionId!, outputPath),
            "close" => CloseDocument(sessionId!, discard),
            "list" => ListSessions(),
            "status" => GetStatus(sessionId!),
            // Temp file operations
            "list_temp" => ListTempFiles(),
            "recover" => RecoverTempFile(sessionId!, outputPath, deleteAfterRecover),
            "delete_temp" => DeleteTempFile(sessionId!),
            "cleanup" => CleanupTempFiles(),
            "temp_stats" => GetTempStats(),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Opens a document and creates a new session.
    /// </summary>
    /// <param name="path">The file path of the document to open.</param>
    /// <param name="mode">The access mode ('readonly' or 'readwrite').</param>
    /// <returns>A JSON string containing the session information.</returns>
    /// <exception cref="ArgumentException">Thrown when path is null or empty.</exception>
    private string OpenDocument(string path, string mode)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("path is required for 'open' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        var newSessionId = _sessionManager.OpenDocument(path, identity, mode);
        var session = _sessionManager.GetSessionStatus(newSessionId, identity);

        return JsonSerializer.Serialize(new
        {
            success = true,
            sessionId = newSessionId,
            message = "Document opened successfully",
            session
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Saves the document in the specified session to file.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <param name="outputPath">Optional output path. If null, saves to original path.</param>
    /// <returns>A JSON string containing the save result.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    private string SaveDocument(string sessionId, string? outputPath)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'save' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        _sessionManager.SaveDocument(sessionId, identity, outputPath);
        var session = _sessionManager.GetSessionStatus(sessionId, identity);

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = outputPath != null
                ? $"Document saved to: {outputPath}"
                : "Document saved to original path",
            session
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Closes a document session.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <param name="discard">If true, discards unsaved changes; otherwise saves before closing.</param>
    /// <returns>A JSON string containing the close result.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    private string CloseDocument(string sessionId, bool discard)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'close' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        _sessionManager.CloseDocument(sessionId, identity, discard);

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = discard
                ? "Session closed (changes discarded)"
                : "Session closed (changes saved)"
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Lists all active document sessions.
    /// </summary>
    /// <returns>A JSON string containing the list of sessions and memory usage.</returns>
    private string ListSessions()
    {
        var identity = _identityAccessor.GetCurrentIdentity();
        var sessions = _sessionManager.ListSessions(identity).ToList();

        return JsonSerializer.Serialize(new
        {
            success = true,
            count = sessions.Count,
            totalMemoryMB = _sessionManager.GetTotalMemoryMb(),
            sessions
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Gets the status of a specific document session.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <returns>A JSON string containing the session status.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    /// <exception cref="KeyNotFoundException">Thrown when the session is not found.</exception>
    private string GetStatus(string sessionId)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'status' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        var session = _sessionManager.GetSessionStatus(sessionId, identity);

        if (session == null)
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        return JsonSerializer.Serialize(new
        {
            success = true,
            session
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Lists all recoverable temporary files.
    /// </summary>
    /// <returns>A JSON string containing the list of recoverable files.</returns>
    private string ListTempFiles()
    {
        var identity = _identityAccessor.GetCurrentIdentity();
        var files = _tempFileManager.ListRecoverableFiles(identity).ToList();

        return JsonSerializer.Serialize(new
        {
            success = true,
            count = files.Count,
            files = files.Select(f => new
            {
                f.SessionId,
                f.OriginalPath,
                f.DocumentType,
                f.SavedAt,
                f.ExpiresAt,
                FileSizeMb = f.FileSizeBytes / (1024.0 * 1024.0),
                f.PromptOnReconnect
            })
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Recovers a temporary file to the specified path.
    /// </summary>
    /// <param name="sessionId">The session identifier to recover.</param>
    /// <param name="outputPath">Optional output path. If null, recovers to original path.</param>
    /// <param name="deleteAfterRecover">Whether to delete the temp file after recovery.</param>
    /// <returns>A JSON string containing the recovery result.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    private string RecoverTempFile(string sessionId, string? outputPath, bool deleteAfterRecover)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'recover' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        var result = _tempFileManager.RecoverSession(sessionId, identity, outputPath, deleteAfterRecover);

        return JsonSerializer.Serialize(new
        {
            result.Success,
            result.SessionId,
            result.RecoveredPath,
            result.OriginalPath,
            result.ErrorMessage,
            message = result.Success
                ? $"Successfully recovered to: {result.RecoveredPath}"
                : $"Recovery failed: {result.ErrorMessage}"
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Deletes a specific temporary file.
    /// </summary>
    /// <param name="sessionId">The session identifier to delete.</param>
    /// <returns>A JSON string containing the deletion result.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    private string DeleteTempFile(string sessionId)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'delete_temp' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        var deleted = _tempFileManager.DeleteTempSession(sessionId, identity);

        return JsonSerializer.Serialize(new
        {
            success = deleted,
            sessionId,
            message = deleted
                ? "Temp file deleted successfully"
                : "Temp file not found"
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Cleans up expired temporary files.
    /// </summary>
    /// <returns>A JSON string containing the cleanup result.</returns>
    private string CleanupTempFiles()
    {
        var result = _tempFileManager.CleanupExpiredFiles();

        return JsonSerializer.Serialize(new
        {
            success = true,
            result.ScannedCount,
            result.DeletedCount,
            result.ErrorCount,
            message = $"Cleaned up {result.DeletedCount} expired files"
        }, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Gets temp file statistics.
    /// </summary>
    /// <returns>A JSON string containing temp file statistics.</returns>
    private string GetTempStats()
    {
        var stats = _tempFileManager.GetStats();

        return JsonSerializer.Serialize(new
        {
            success = true,
            stats.TotalCount,
            stats.TotalSizeMb,
            stats.ExpiredCount,
            retentionHours = _sessionManager.Config.TempRetentionHours
        }, new JsonSerializerOptions { WriteIndented = true });
    }
}