using System.ComponentModel;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Results.Session;
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
    /// <returns>Operation result object</returns>
    [McpServerTool(
        Name = "document_session",
        Title = "Manage Document Sessions",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [OutputSchema(typeof(DocumentSessionResults))]
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
    public object Execute(
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
        object result = operation.ToLower() switch
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
        return ResultHelper.FinalizeResult((dynamic)result, outputPath, sessionId);
    }

    /// <summary>
    ///     Opens a document and creates a new session.
    /// </summary>
    /// <param name="path">The file path of the document to open.</param>
    /// <param name="mode">The access mode ('readonly' or 'readwrite').</param>
    /// <returns>The open session result.</returns>
    /// <exception cref="ArgumentException">Thrown when path is null or empty.</exception>
    private OpenSessionResult OpenDocument(string path, string mode)
    {
        if (string.IsNullOrEmpty(path))
            throw new ArgumentException("path is required for 'open' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        var newSessionId = _sessionManager.OpenDocument(path, identity, mode);
        var session = _sessionManager.GetSessionStatus(newSessionId, identity);

        return new OpenSessionResult
        {
            Success = true,
            SessionId = newSessionId,
            Message = "Document opened successfully",
            Session = ToSessionInfoDto(session!)
        };
    }

    /// <summary>
    ///     Saves the document in the specified session to file.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <param name="outputPath">Optional output path. If null, saves to original path.</param>
    /// <returns>The save session result.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    private SaveSessionResult SaveDocument(string sessionId, string? outputPath)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'save' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        _sessionManager.SaveDocument(sessionId, identity, outputPath);
        var session = _sessionManager.GetSessionStatus(sessionId, identity);

        return new SaveSessionResult
        {
            Success = true,
            Message = outputPath != null
                ? $"Document saved to: {outputPath}"
                : "Document saved to original path",
            Session = ToSessionInfoDto(session!)
        };
    }

    /// <summary>
    ///     Closes a document session.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <param name="discard">If true, discards unsaved changes; otherwise saves before closing.</param>
    /// <returns>The close session result.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    private CloseSessionResult CloseDocument(string sessionId, bool discard)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'close' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        _sessionManager.CloseDocument(sessionId, identity, discard);

        return new CloseSessionResult
        {
            Success = true,
            Message = discard
                ? "Session closed (changes discarded)"
                : "Session closed (changes saved)"
        };
    }

    /// <summary>
    ///     Lists all active document sessions.
    /// </summary>
    /// <returns>The list sessions result.</returns>
    private ListSessionsResult ListSessions()
    {
        var identity = _identityAccessor.GetCurrentIdentity();
        var sessions = _sessionManager.ListSessions(identity).ToList();

        return new ListSessionsResult
        {
            Success = true,
            Count = sessions.Count,
            TotalMemoryMb = _sessionManager.GetTotalMemoryMb(),
            Sessions = sessions.Select(ToSessionInfoDto).ToList()
        };
    }

    /// <summary>
    ///     Gets the status of a specific document session.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <returns>The session status result.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    /// <exception cref="KeyNotFoundException">Thrown when the session is not found.</exception>
    private SessionStatusResult GetStatus(string sessionId)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'status' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        var session = _sessionManager.GetSessionStatus(sessionId, identity);

        if (session == null)
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        return new SessionStatusResult
        {
            Success = true,
            Session = ToSessionInfoDto(session)
        };
    }

    /// <summary>
    ///     Lists all recoverable temporary files.
    /// </summary>
    /// <returns>The list temp files result.</returns>
    private ListTempFilesResult ListTempFiles()
    {
        var identity = _identityAccessor.GetCurrentIdentity();
        var files = _tempFileManager.ListRecoverableFiles(identity).ToList();

        return new ListTempFilesResult
        {
            Success = true,
            Count = files.Count,
            Files = files.Select(f => new TempFileInfoDto
            {
                SessionId = f.SessionId,
                OriginalPath = f.OriginalPath,
                DocumentType = f.DocumentType,
                SavedAt = f.SavedAt,
                ExpiresAt = f.ExpiresAt,
                FileSizeMb = f.FileSizeBytes / (1024.0 * 1024.0),
                PromptOnReconnect = f.PromptOnReconnect
            }).ToList()
        };
    }

    /// <summary>
    ///     Recovers a temporary file to the specified path.
    /// </summary>
    /// <param name="sessionId">The session identifier to recover.</param>
    /// <param name="outputPath">Optional output path. If null, recovers to original path.</param>
    /// <param name="deleteAfterRecover">Whether to delete the temp file after recovery.</param>
    /// <returns>The recover temp file result.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    private RecoverTempFileResult RecoverTempFile(string sessionId, string? outputPath, bool deleteAfterRecover)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'recover' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        var result = _tempFileManager.RecoverSession(sessionId, identity, outputPath, deleteAfterRecover);

        return new RecoverTempFileResult
        {
            Success = result.Success,
            SessionId = result.SessionId,
            RecoveredPath = result.RecoveredPath,
            OriginalPath = result.OriginalPath,
            ErrorMessage = result.ErrorMessage,
            Message = result.Success
                ? $"Successfully recovered to: {result.RecoveredPath}"
                : $"Recovery failed: {result.ErrorMessage}"
        };
    }

    /// <summary>
    ///     Deletes a specific temporary file.
    /// </summary>
    /// <param name="sessionId">The session identifier to delete.</param>
    /// <returns>The delete temp file result.</returns>
    /// <exception cref="ArgumentException">Thrown when sessionId is null or empty.</exception>
    private DeleteTempFileResult DeleteTempFile(string sessionId)
    {
        if (string.IsNullOrEmpty(sessionId))
            throw new ArgumentException("sessionId is required for 'delete_temp' operation");

        var identity = _identityAccessor.GetCurrentIdentity();
        var deleted = _tempFileManager.DeleteTempSession(sessionId, identity);

        return new DeleteTempFileResult
        {
            Success = deleted,
            SessionId = sessionId,
            Message = deleted
                ? "Temp file deleted successfully"
                : "Temp file not found"
        };
    }

    /// <summary>
    ///     Cleans up expired temporary files.
    /// </summary>
    /// <returns>The cleanup temp files result.</returns>
    private CleanupTempFilesResult CleanupTempFiles()
    {
        var result = _tempFileManager.CleanupExpiredFiles();

        return new CleanupTempFilesResult
        {
            Success = true,
            ScannedCount = result.ScannedCount,
            DeletedCount = result.DeletedCount,
            ErrorCount = result.ErrorCount,
            Message = $"Cleaned up {result.DeletedCount} expired files"
        };
    }

    /// <summary>
    ///     Gets temp file statistics.
    /// </summary>
    /// <returns>The temp file stats result.</returns>
    private TempFileStatsResult GetTempStats()
    {
        var stats = _tempFileManager.GetStats();

        return new TempFileStatsResult
        {
            Success = true,
            TotalCount = stats.TotalCount,
            TotalSizeMb = stats.TotalSizeMb,
            ExpiredCount = stats.ExpiredCount,
            RetentionHours = _sessionManager.Config.TempRetentionHours
        };
    }

    /// <summary>
    ///     Converts a SessionInfo to SessionInfoDto.
    /// </summary>
    /// <param name="session">The session info to convert.</param>
    /// <returns>The session info DTO.</returns>
    private static SessionInfoDto ToSessionInfoDto(SessionInfo session)
    {
        return new SessionInfoDto
        {
            SessionId = session.SessionId,
            DocumentType = session.DocumentType,
            Path = session.Path,
            Mode = session.Mode,
            IsDirty = session.IsDirty,
            OpenedAt = session.OpenedAt,
            LastAccessedAt = session.LastAccessedAt,
            EstimatedMemoryMb = session.EstimatedMemoryMb
        };
    }
}
