using System.Collections.Concurrent;
using System.Diagnostics.CodeAnalysis;
using System.Security.Cryptography;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Slides;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Manages document sessions for in-memory document editing
/// </summary>
public class DocumentSessionManager : IDisposable
{
    /// <summary>
    ///     Optional server configuration. When non-null, every user-supplied path entering the
    ///     session manager is checked against <see cref="ServerConfig.AllowedBasePaths" /> in
    ///     addition to the standard <see cref="SecurityHelper.ValidateFilePath" /> shape check.
    ///     Test constructions and DI consumers that don't supply this leave it null →
    ///     allowlist becomes a no-op (backward compatible).
    /// </summary>
    /// <summary>
    ///     How long auto-save waits for a session's in-flight operations before skipping it for
    ///     this tick. Kept short so one busy session cannot stall the whole auto-save pass.
    /// </summary>
    private const int AutoSaveDrainTimeoutMs = 2_000;

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

    private readonly ServerConfig? _serverConfig;

    /// <summary>
    ///     Thread-safe dictionary of active sessions grouped by owner key
    ///     Key: Owner storage key, Value: Dictionary of SessionId -> Session
    /// </summary>
    private readonly ConcurrentDictionary<string, ConcurrentDictionary<string, DocumentSession>> _sessionsByOwner =
        new();

    /// <summary>
    ///     Re-entrancy guard for <see cref="AutoSaveDirtySessions" /> (1 = a callback is running)
    /// </summary>
    private int _autoSaveRunning;

    /// <summary>
    ///     Re-entrancy guard for <see cref="CleanupIdleSessions" /> (1 = a callback is running)
    /// </summary>
    private int _cleanupRunning;

    /// <summary>
    ///     Tracks whether this manager has been disposed (0 = not disposed, 1 = disposed)
    /// </summary>
    private int _disposed;

    /// <summary>
    ///     Creates a new document session manager
    /// </summary>
    /// <param name="config">Session configuration</param>
    /// <param name="loggerFactory">Logger factory for logging</param>
    /// <param name="serverConfig">
    ///     Optional server configuration carrying the
    ///     <see cref="ServerConfig.AllowedBasePaths" /> allowlist. When null (default,
    ///     used by tests), allowlist enforcement is skipped while shape validation
    ///     (<see cref="SecurityHelper.ValidateFilePath" />) still runs on every
    ///     user-supplied path.
    /// </param>
    public DocumentSessionManager(SessionConfig config, ILoggerFactory? loggerFactory = null,
        ServerConfig? serverConfig = null)
    {
        Config = config;
        _logger = loggerFactory?.CreateLogger<DocumentSessionManager>();
        _serverConfig = serverConfig;

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
    ///     Invoked once a close has taken the session exclusively, before the final save and
    ///     dispose.
    ///     <para>
    ///         That span is where a request holding a reference from before the close used to be
    ///         able to acquire the session (R4-S06). A fixture cannot land inside it by racing,
    ///         so this lets one act there every time. Unset in production, where it costs a null
    ///         check.
    ///     </para>
    /// </summary>
    internal Action<DocumentSession>? OnSessionSealed { get; set; }

    /// <summary>
    ///     How long a close waits for the operations already running before it seals the session
    ///     without them.
    ///     <para>
    ///         The production value is the session's own drain timeout. A fixture lowers it so the
    ///         whole ordered sequence — an operation in flight, a close that unregisters the session
    ///         and hands the document on, then that operation finishing last — can be driven in
    ///         milliseconds instead of half a minute (R4-S06).
    ///     </para>
    /// </summary>
    internal int ClosingDrainTimeoutMs { get; set; } = DocumentSession.DrainTimeoutMs;

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
        if (Interlocked.Exchange(ref _disposed, 1) == 1)
            return;

        _cleanupTimer?.Dispose();
        _autoSaveTimer?.Dispose();

        // Disposal must apply the same disconnect policy as an orderly shutdown. Disposing the
        // sessions directly would drop unsaved changes whenever disposal is reached by a path
        // other than SessionLifetimeService, which is the case for a host that faults during
        // startup or a test that disposes the manager itself. OnServerShutdown is idempotent:
        // it drains, saves or discards per configuration, disposes, and clears the registry.
        try
        {
            OnServerShutdown();
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error applying disconnect policy while disposing the session manager");

            foreach (var ownerSessions in _sessionsByOwner.Values)
            foreach (var session in ownerSessions.Values)
                session.Dispose();

            _sessionsByOwner.Clear();
        }
    }

    /// <summary>
    ///     Event raised when a session's IsDirty property becomes true.
    ///     Parameters: sessionId, requestor identity.
    /// </summary>
    public event Action<string, SessionIdentity>? SessionModified;

    /// <summary>
    ///     Event raised when a session is closed.
    ///     Parameters: sessionId, owner identity.
    /// </summary>
    [SuppressMessage("Major Code Smell", "S3264:Events should be invoked",
        Justification =
            "NotifySessionClosed invokes each subscriber from GetInvocationList so one failing subscriber cannot block the others.")]
    public event Action<string, SessionIdentity>? SessionClosed;

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

        // Validate BEFORE new FileInfo(path) to close the existence/size stat-oracle
        // side channel (bug 20260415-session-loader-path, HIGH-3) and the four
        // LoadDocument read sinks (Word/Excel/PPT/PDF) at once. Single upstream guard
        // mirrors the Core/Session/DocumentContext.cs:147-149 in-repo precedent.
        // The loader below opens the canonical path, not the one the caller wrote, so a link
        // swapped between the check and the open cannot redirect it (R3-S03).
        path = SecurityHelper.ValidateUserPath(path, _serverConfig?.AllowedBasePaths ?? []);

        var ownerKey = owner.GetStorageKey(Config.IsolationMode);
        var ownerSessions =
            _sessionsByOwner.GetOrAdd(ownerKey, _ => new ConcurrentDictionary<string, DocumentSession>());

        if (ownerSessions.Count >= Config.MaxSessions)
            throw new InvalidOperationException($"Maximum session limit ({Config.MaxSessions}) reached for this user");

        var fileInfo = new FileInfo(path);
        if (!fileInfo.Exists)
        {
            _logger?.LogWarning("OpenDocument: file not found at {Path}", path);
            throw new FileNotFoundException("File not found");
        }

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
                if (session.IsDisposed)
                {
                    _logger?.LogWarning("Attempted to access disposed session {SessionId}", sessionId);
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
        if (session != null)
        {
            session.IsDirty = true;
            SessionModified?.Invoke(sessionId, requestor);
        }
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
        // Validate the user-supplied outputPath at the trust boundary, BEFORE looking up
        // the session, so an invalid path never reaches the .Save sink (bug
        // 20260415-session-loader-path, U2). Null/empty means "use session.Path", which
        // was validated when the session was opened.
        if (!string.IsNullOrEmpty(outputPath))
            outputPath = SecurityHelper.ValidateUserPath(outputPath,
                _serverConfig?.AllowedBasePaths ?? [], nameof(outputPath));

        var session = FindSessionWithAuth(sessionId, requestor);
        if (session == null)
            throw new KeyNotFoundException($"Session not found: {sessionId}");

        if (session.Mode == "readonly") throw new InvalidOperationException("Cannot save a readonly session");

        var savePath = outputPath ?? session.Path;

        // Re-check allowlist on the resolved savePath right before write (HIGH-1).
        // session.Path was admitted at open time, but ServerConfig.AllowedBasePaths is
        // mutable (Core/ServerConfig.cs:226-241). A configuration reload that narrows
        // the allowlist must not let pre-existing sessions auto-save outside the new
        // bounds. Path-shape is immutable so ValidateFilePath is not re-run.
        savePath = ReassertAllowlistForResolvedPath(savePath, nameof(savePath));

        // A handler mutates the document outside the session lock (it holds a usage scope instead),
        // so the lock alone does not stop a save from capturing a half-applied change. Waiting for
        // in-flight operations to finish makes the saved file a complete state, never a torn one.
        // BeginExclusive rather than a bare drain: waiting for the count to reach zero says
        // nothing unless new operations are refused while the save runs (R2-S09).
        if (!session.BeginExclusive())
            throw new InvalidOperationException(
                $"Session {sessionId} is busy with another operation and cannot be saved right now.");

        try
        {
            session.Execute(doc => SaveDocumentToFile(doc, session.Type, savePath));
            session.IsDirty = false;
        }
        finally
        {
            // The session keeps serving requests after an explicit save. The exception is a close
            // that gave up waiting for this save: it leaves the document here rather than disposing
            // it underneath the write (R4-S06), and the session is unregistered by then, so
            // releasing it is this caller's job — nothing else can reach it.
            // EndExclusive settles this itself now, so there is one place that decides rather
            // than a copy of the rule at every release site (R8-S03).
            session.EndExclusive();
        }

        _logger?.LogInformation("Saved session {SessionId} to {Path}", sessionId, savePath);
    }

    /// <summary>
    ///     Closes a session (no authorization check)
    /// </summary>
    /// <param name="sessionId">Session ID to close</param>
    /// <param name="discard">If true, discard unsaved changes; otherwise auto-save</param>
    /// <exception cref="KeyNotFoundException">Thrown when session not found</exception>
    /// <remarks>
    ///     The session is taken exclusively before anything is saved, and is never released: once
    ///     a close begins, a reference obtained earlier can no longer be used (R4-S06). If
    ///     in-flight operations do not finish within the drain timeout the document is not saved,
    ///     because it may be mid-mutation; the close still completes and the session still goes
    ///     away.
    /// </remarks>
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

        if (ownerDict != null && ownerKey != null && ownerDict.IsEmpty)
            ((ICollection<KeyValuePair<string, ConcurrentDictionary<string, DocumentSession>>>)_sessionsByOwner)
                .Remove(new KeyValuePair<string, ConcurrentDictionary<string, DocumentSession>>(ownerKey, ownerDict));

        // Notified after the session is closed, not before. The session had already been taken
        // out of the registry by this point, so a subscriber that threw skipped the seal, the save
        // and the dispose below and left a document nothing could reach and nothing would release
        // (R8-S01). Each subscriber is now isolated as well, so one failing observer cannot stop
        // the others being told either.
        //
        // Unregistering stops new lookups, but a request that already holds a reference can be
        // sitting between GetSession() and AcquireUsage(): draining alone let it acquire the
        // moment the count reached zero and then race the save and dispose below (R4-S06). The
        // barrier goes up first and is never released, so nothing can acquire again.
        var outcome = session.BeginClosing(ClosingDrainTimeoutMs);
        OnSessionSealed?.Invoke(session);
        LogSealOutcome(sessionId, outcome);

        // Only an outright seal means the document is this caller's to write and to release. A
        // save still holding it will finish and release it itself; disposing it here is what
        // failed that save and lost both copies of the work (R4-S06).
        if (outcome != SessionSealOutcome.Sealed)
        {
            // The document belongs to whoever is still using it; they release it when they finish.
            _logger?.LogInformation("Closed session {SessionId} (discard={Discard})", sessionId, discard);
            NotifySessionClosed(sessionId, session.Owner);
            return;
        }

        var sessionType = session.Type;
        var sessionPath = session.Path;
        try
        {
            if (!discard && session.IsDirty)
            {
                sessionPath = ReassertAllowlistForResolvedPath(sessionPath, nameof(sessionPath));
                session.Execute(doc => SaveDocumentToFile(doc, sessionType, sessionPath));
            }
        }
        finally
        {
            // Sealed means this close is the last one out, but it still asks: an exclusive holder
            // or a usage scope must never dispose the same document (R4-S06).
            if (session.TryClaimDisposeOwnership()) session.Dispose();
            NotifySessionClosed(sessionId, session.Owner);
        }

        _logger?.LogInformation("Closed session {SessionId} (discard={Discard})", sessionId, discard);
    }

    /// <summary>
    ///     Tells every <see cref="SessionClosed" /> subscriber, one at a time, and never lets one
    ///     of them affect the caller.
    /// </summary>
    /// <param name="sessionId">The session that closed.</param>
    /// <param name="owner">Who owned it.</param>
    private void NotifySessionClosed(string sessionId, SessionIdentity owner)
    {
        foreach (var subscriber in SessionClosed?.GetInvocationList() ?? [])
            try
            {
                ((Action<string, SessionIdentity>)subscriber)(sessionId, owner);
            }
            catch (Exception error)
            {
                _logger?.LogWarning(error,
                    "A SessionClosed subscriber failed for session {SessionId}; the session was "
                    + "closed regardless", sessionId);
            }
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
                // Only the caller that actually removed the session may disconnect it, so a
                // concurrent close/cleanup cannot double-handle the same session.
                if (RemoveSessionById(session.SessionId))
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
    /// <returns><c>true</c> when this call removed the session; <c>false</c> when it was already gone.</returns>
    private bool RemoveSessionById(string sessionId)
    {
        foreach (var kvp in _sessionsByOwner)
            if (kvp.Value.TryRemove(sessionId, out _))
            {
                if (kvp.Value.IsEmpty)
                    ((ICollection<KeyValuePair<string, ConcurrentDictionary<string, DocumentSession>>>)_sessionsByOwner)
                        .Remove(new KeyValuePair<string, ConcurrentDictionary<string, DocumentSession>>(kvp.Key,
                            kvp.Value));
                return true;
            }

        return false;
    }

    /// <summary>
    ///     Re-asserts the admin-configured path allowlist against a save path immediately before the
    ///     .Save sink, resolving symbolic links in the process. Closes the configuration-reload race
    ///     (bug 20260415-session-loader-path, HIGH-1) and the symlink TOCTOU gap
    ///     (bug 20260415-symlink-toctou-sweep, Phase 1): a session path that passed the allowlist
    ///     check at open time could be a symlink whose target escapes the (possibly narrowed) current
    ///     allowlist. Path shape validation (<see cref="SecurityHelper.ValidateFilePath" />) is not
    ///     re-invoked here because shape is immutable for the session lifetime. No-op when no
    ///     <c>ServerConfig</c> is wired in (test contexts).
    /// </summary>
    /// <param name="resolvedPath">
    ///     The write target (typically <c>session.Path</c> or a generated temp path under
    ///     <c>Config.TempDirectory</c>).
    /// </param>
    /// <param name="paramName">Parameter name used in thrown exception messages.</param>
    /// <returns>
    ///     The canonical path. Callers must write to this rather than to the path they passed in:
    ///     the two differ exactly when something was resolved, which is the case the check exists
    ///     for (R3-S03).
    /// </returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the path (after following symlinks) is outside the configured allowlist,
    ///     or when a circular symbolic link is detected.
    /// </exception>
    private string ReassertAllowlistForResolvedPath(string resolvedPath, string paramName)
    {
        var allowed = _serverConfig?.AllowedBasePaths;
        return allowed is { Count: > 0 }
            ? SecurityHelper.ResolveAndEnsureWithinAllowlist(resolvedPath, allowed, paramName)
            : resolvedPath;
    }

    /// <summary>
    ///     Records why a close could not take the document, if it could not.
    /// </summary>
    /// <param name="sessionId">The session being closed.</param>
    /// <param name="outcome">What sealing it achieved.</param>
    private void LogSealOutcome(string sessionId, SessionSealOutcome outcome)
    {
        switch (outcome)
        {
            case SessionSealOutcome.HeldByAnotherOperation:
                _logger?.LogError(
                    "Session {SessionId} was held by another save after the timeout; that operation " +
                    "owns the document and this close neither wrote nor released it",
                    sessionId);
                break;
            case SessionSealOutcome.ActiveOperationsRemain:
                _logger?.LogError(
                    "Session {SessionId} still had in-flight operations after the drain timeout; " +
                    "unsaved changes were not written because the document may be mid-mutation",
                    sessionId);
                break;
            case SessionSealOutcome.AlreadyClosing:
                _logger?.LogWarning(
                    "Session {SessionId} was already being closed by another caller", sessionId);
                break;
            case SessionSealOutcome.Sealed:
            default:
                break;
        }
    }

    /// <summary>
    ///     Handles disconnect behavior for a session based on configuration.
    ///     Disposes the session once it is sealed, whether or not the save succeeded and whether or
    ///     not it was dirty. A session another operation still holds is left to that operation to
    ///     release, because disposing it there fails the save that is writing it (R4-S06).
    /// </summary>
    /// <param name="session">Session to handle disconnect for</param>
    private void HandleDisconnect(DocumentSession session)
    {
        // Same rule on this path: the notification happens once the session has actually been
        // dealt with, so a throwing subscriber cannot skip the seal (R8-S01).
        //
        // Sealing, not borrowing the save barrier: a disconnect handled while an explicit save
        // held the session reported a drain timeout it had not observed, then disposed the
        // document under that save (R4-S06). A save in flight is waited for, and if it outlasts
        // the timeout the document is left to it rather than taken away.
        var outcome = session.BeginClosing(ClosingDrainTimeoutMs);
        LogSealOutcome(session.SessionId, outcome);
        if (outcome != SessionSealOutcome.Sealed)
        {
            NotifySessionClosed(session.SessionId, session.Owner);
            return;
        }

        var sessionType = session.Type;
        var sessionPath = session.Path;
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
                    sessionPath = ReassertAllowlistForResolvedPath(sessionPath, nameof(sessionPath));
                    session.Execute(doc => SaveDocumentToFile(doc, sessionType, sessionPath));
                    DeleteSessionTempFiles(session.SessionId);
                    _logger?.LogInformation("Auto-saved session {SessionId} and cleaned up temp files",
                        session.SessionId);
                    break;

                case DisconnectBehavior.SaveToTemp:
                    var tempPath = GetTempPath(session);
                    tempPath = ReassertAllowlistForResolvedPath(tempPath, nameof(tempPath));
                    session.Execute(doc => SaveDocumentToFile(doc, sessionType, tempPath));
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
                    promptTempPath = ReassertAllowlistForResolvedPath(promptTempPath, nameof(promptTempPath));
                    session.Execute(doc => SaveDocumentToFile(doc, sessionType, promptTempPath));
                    SaveSessionMetadata(session, promptTempPath, true);
                    _logger?.LogInformation("Saved session {SessionId} for prompt on reconnect", session.SessionId);
                    break;
            }
        }
        finally
        {
            if (session.TryClaimDisposeOwnership()) session.Dispose();
            NotifySessionClosed(session.SessionId, session.Owner);
        }
    }

    /// <summary>
    ///     Timer callback to cleanup idle sessions.
    ///     Non-reentrant: a slow save in one tick must not overlap the next tick's callback.
    /// </summary>
    /// <param name="state">Timer state (not used)</param>
    private void CleanupIdleSessions(object? state)
    {
        if (Interlocked.Exchange(ref _cleanupRunning, 1) == 1)
            return;

        try
        {
            var timeout = TimeSpan.FromMinutes(Config.IdleTimeoutMinutes);
            var now = DateTime.UtcNow;

            var allSessions = _sessionsByOwner.Values
                .SelectMany(s => s.Values)
                .ToList();

            foreach (var session in allSessions)
                if (now - session.LastAccessedAt > timeout)
                {
                    if (session.HasActiveUsers)
                    {
                        _logger?.LogDebug(
                            "Skipping cleanup of idle session {SessionId} because it has active users",
                            session.SessionId);
                        continue;
                    }

                    _logger?.LogInformation("Session {SessionId} timed out after {Minutes} minutes of inactivity",
                        session.SessionId, Config.IdleTimeoutMinutes);

                    try
                    {
                        // Only the caller that actually removed the session may disconnect it,
                        // so a concurrent close cannot double-handle the same session.
                        if (RemoveSessionById(session.SessionId))
                            HandleDisconnect(session);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, "Error cleaning up idle session {SessionId}", session.SessionId);
                    }
                }
        }
        finally
        {
            Volatile.Write(ref _cleanupRunning, 0);
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
        // Non-reentrant: a slow save in one tick must not overlap the next tick's callback.
        if (Interlocked.Exchange(ref _autoSaveRunning, 1) == 1)
            return;

        try
        {
            var dirtySessions = _sessionsByOwner.Values
                .SelectMany(s => s.Values)
                .Where(s => s is { IsDirty: true, IsDisposed: false })
                .ToList();

            if (dirtySessions.Count == 0)
                return;

            _logger?.LogDebug("Auto-saving {Count} dirty sessions", dirtySessions.Count);

            foreach (var session in dirtySessions)
            {
                // A busy session is skipped rather than captured mid-mutation; the next tick
                // retries, and the session stays dirty until one of them succeeds. The acquisition
                // is outside the try so the finally below can only ever release a lock this loop
                // actually took: releasing after a refused acquisition tore down whichever save or
                // close did hold it (R3-S04).
                if (!session.BeginExclusive(AutoSaveDrainTimeoutMs))
                {
                    _logger?.LogDebug("Skipped auto-save of busy session {SessionId}", session.SessionId);
                    continue;
                }

                try
                {
                    var tempPath = GetTempPath(session);
                    tempPath = ReassertAllowlistForResolvedPath(tempPath, nameof(tempPath));

                    // Write beside the slot and swap, so a failure part-way through leaves the
                    // previous recovery file intact instead of a truncated one.
                    // Resolved in its own right: it is a different leaf from the temp path
                    // resolved above, so a symlink planted at this name was never checked and the
                    // library's save followed it out of the allowlist (R4-S07).
                    var stagingPath = ReassertAllowlistForResolvedPath(
                        BuildStagingPath(tempPath), "stagingPath");
                    session.Execute(doc => SaveDocumentToFile(doc, session.Type, stagingPath));
                    File.Move(stagingPath, tempPath, true);
                    SaveSessionMetadata(session, tempPath);
                    _logger?.LogInformation("Auto-saved dirty session {SessionId} to temp: {TempPath}",
                        session.SessionId, tempPath);
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "Error auto-saving session {SessionId}", session.SessionId);
                }
                finally
                {
                    // The session goes on serving requests after an auto-save, so the barrier this
                    // loop took must come down on every path out — including the failure one. A
                    // close that gave up waiting for this auto-save left the document here, so
                    // releasing it is then this loop's job (R4-S06).
                    session.EndExclusive();
                }
            }
        }
        finally
        {
            Volatile.Write(ref _autoSaveRunning, 0);
        }
    }

    /// <summary>
    ///     Builds the path auto-save writes to before swapping it over the recovery slot.
    ///     <para>
    ///         The suffix goes before the extension, not after it. Appending <c>.new</c> to
    ///         <c>session.pptx</c> produced <c>session.pptx.new</c>, and the save format is
    ///         resolved from the extension: <see cref="Helpers.PowerPoint.PptSaveFormatResolver" />
    ///         refuses <c>.new</c> outright, so every periodic auto-save of a PowerPoint session
    ///         threw and was swallowed by the surrounding catch, leaving nothing but a log line
    ///         (R2-C01). Keeping the real extension keeps the staging file the same format as the
    ///         file it replaces.
    ///     </para>
    /// </summary>
    /// <param name="finalPath">Path the staging file will be moved onto.</param>
    /// <returns>The staging path, carrying the same extension as <paramref name="finalPath" />.</returns>
    internal static string BuildStagingPath(string finalPath)
    {
        var directory = Path.GetDirectoryName(finalPath);
        var stem = Path.GetFileNameWithoutExtension(finalPath);
        var extension = Path.GetExtension(finalPath);
        var name = $"{stem}.new{extension}";
        return string.IsNullOrEmpty(directory) ? name : Path.Combine(directory, name);
    }

    /// <summary>
    ///     Generates a unique session ID
    /// </summary>
    /// <returns>Generated session ID in format sess_XXXXXXXXXXXXXXXXXXXXXXXX (24 chars total)</returns>
    private static string GenerateSessionId()
    {
        return $"sess_{Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLowerInvariant()}";
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
    private object LoadDocument(string path, DocumentType type)
    {
        return type switch
        {
            DocumentType.Word => GuardedWordLoader.Load(path, _serverConfig?.AllowedBasePaths ?? []),
            DocumentType.Excel => new Workbook(path),
            DocumentType.PowerPoint => LoadPresentation(path),
            DocumentType.Pdf => new Document(path),
            _ => throw new NotSupportedException($"Unsupported document type: {type}")
        };
    }

    /// <summary>
    ///     Opens a presentation with this process's Aspose.Slides serialisation in force.
    ///     <para>
    ///         The library raised a null-dereference from its own loading path when another thread
    ///         was inside it; the session's own lock does not help because it is per session and
    ///         this load happens before there is one (SlidesGate).
    ///     </para>
    /// </summary>
    /// <param name="path">The presentation to open.</param>
    /// <returns>The loaded presentation.</returns>
    private static Presentation LoadPresentation(string path)
    {
        using var gate = SlidesGate.Enter();
        return new Presentation(path);
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
                ((Aspose.Words.Document)document).Save(path);
                break;
            case DocumentType.Excel:
                ((Workbook)document).Save(path);
                break;
            case DocumentType.PowerPoint:
                // Every save of a presentation funnels through here — explicit save, save on
                // close, all three disconnect behaviours and autosave — so this one gate covers
                // five call sites that each had nothing between them and the library (SlidesGate).
                using (SlidesGate.Enter())
                {
                    ((Presentation)document).Save(path, PptSaveFormatResolver.Resolve(path));
                }

                break;
            case DocumentType.Pdf:
                ((Document)document).Save(path);
                break;
        }
    }

    /// <summary>
    ///     Returns the recovery file path for a session. Every auto-save of the same session writes
    ///     the same slot, so a long-lived session keeps exactly one recovery file instead of one per
    ///     tick, and <c>list_temp_files</c> shows one entry rather than a growing history.
    /// </summary>
    /// <param name="session">Session to generate temp path for</param>
    /// <returns>Full path to temporary file</returns>
    private string GetTempPath(DocumentSession session)
    {
        var ext = Path.GetExtension(session.Path);
        return Path.Combine(Config.TempDirectory, $"aspose_session_{session.SessionId}_autosave{ext}");
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

                    // T5: symlink-aware check before File.Delete on metadata-sourced TempPath.
                    // ResolveAndEnsureWithinAllowlist follows symbolic links so a planted link
                    // whose lexical path is inside TempDirectory but whose target is outside
                    // cannot be used as a delete-anywhere primitive
                    // (bug 20260415-symlink-toctou-sweep, Phase 1).
                    if (metadata?.TempPath != null && File.Exists(metadata.TempPath))
                        try
                        {
                            SecurityHelper.ResolveAndEnsureWithinAllowlist(
                                metadata.TempPath,
                                [Config.TempDirectory],
                                "tempPath");
                            File.Delete(metadata.TempPath);
                        }
                        catch (ArgumentException ex)
                        {
                            _logger?.LogWarning(
                                ex,
                                "Refusing to delete temp file outside TempDirectory for session {SessionId}",
                                sessionId);
                        }

                    // T6: resolve metadata path before deleting it.
                    SecurityHelper.ResolveAndEnsureWithinAllowlist(
                        metadataPath,
                        [Config.TempDirectory],
                        "metadataPath");
                    File.Delete(metadataPath);
                    _logger?.LogDebug("Deleted temp file for session {SessionId}: {Path}", sessionId, metadataPath);
                }
                catch (ArgumentException ex)
                {
                    _logger?.LogWarning(ex, "Refusing to delete metadata outside TempDirectory: {Path}", metadataPath);
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
        // T7: resolve immediately before File.WriteAllText to catch a symlink planted at
        // metadataPath redirecting the JSON write outside TempDirectory.
        // allowedBase is derived from tempPath's parent since this is a static method without
        // Config access; tempPath was already validated by ReassertAllowlistForResolvedPath
        // at the call site (bug 20260415-symlink-toctou-sweep, Phase 1).
        var metaAllowedBase = Path.GetDirectoryName(Path.GetFullPath(tempPath))
                              ?? throw new InvalidOperationException(
                                  "GetFullPath returned a path without a directory component");
        SecurityHelper.ResolveAndEnsureWithinAllowlist(
            metadataPath,
            [metaAllowedBase],
            "metadataPath");
        File.WriteAllText(metadataPath, JsonSerializer.Serialize(metadata, JsonDefaults.Indented));

        // Both files land in a potentially shared temp directory (default: /tmp on Unix):
        // restrict them to the owner so other local users cannot read document contents.
        SecurityHelper.HardenPrivateFile(tempPath);
        SecurityHelper.HardenPrivateFile(metadataPath);
    }
}
