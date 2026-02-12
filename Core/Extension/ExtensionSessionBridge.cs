using System.Collections.Concurrent;
using AsposeMcpServer.Core.Conversion;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Results.Extension;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Bridges document sessions with extensions, managing bindings and snapshot delivery.
///     Handles session-extension binding lifecycle and coordinated snapshot sending.
/// </summary>
public class ExtensionSessionBridge : IDisposable
{
    /// <summary>
    ///     Separator for binding keys. Uses unit separator character unlikely to appear in IDs.
    /// </summary>
    private const string BindingKeySeparator = "\x1F";

    /// <summary>
    ///     Maximum number of bindings to prevent resource exhaustion.
    /// </summary>
    private const int MaxBindings = 10000;

    /// <summary>
    ///     Maximum number of stale locks to accumulate before triggering cleanup.
    /// </summary>
    /// <remarks>
    ///     Stale locks are semaphores from removed bindings that may still have
    ///     active waiters. They're queued for deferred disposal instead of
    ///     immediate disposal to prevent InvalidOperationException.
    ///     When this limit is exceeded, oldest locks are force-disposed.
    /// </remarks>
    private const int MaxStaleLocks = 100;

    /// <summary>
    ///     Maximum number of concurrent event handling tasks to prevent unbounded growth.
    /// </summary>
    /// <remarks>
    ///     1000 tasks allows for reasonable parallelism while preventing:
    ///     - Memory exhaustion from unbounded task accumulation
    ///     - Thread pool starvation
    ///     When the limit is reached, new events are dropped with a warning log.
    /// </remarks>
    private const int MaxActiveTasks = 1000;

    /// <summary>
    ///     Collection of active event handling tasks for graceful shutdown.
    /// </summary>
    private readonly ConcurrentDictionary<Guid, Task> _activeTasks = new();

    /// <summary>
    ///     Dictionary of per-binding locks to avoid global serialization bottleneck.
    /// </summary>
    private readonly ConcurrentDictionary<string, SemaphoreSlim> _bindingLocks = new();

    /// <summary>
    ///     Dictionary of session-extension bindings keyed by "sessionId{separator}extensionId".
    /// </summary>
    private readonly ConcurrentDictionary<string, SessionBindingInfo> _bindings = new();

    /// <summary>
    ///     Set of recently closed session IDs to prevent stale cache updates.
    ///     Key: sessionId, Value: closure timestamp.
    /// </summary>
    private readonly ConcurrentDictionary<string, DateTime> _closedSessions = new();

    /// <summary>
    ///     Extension configuration containing frame interval and other settings.
    /// </summary>
    private readonly ExtensionConfig _config;

    /// <summary>
    ///     Cache for conversion results to avoid redundant conversions for parallel bindings.
    ///     Key: "sessionId:format", Value: (data, timestamp).
    /// </summary>
    private readonly ConcurrentDictionary<string, (byte[] Data, DateTime Timestamp)> _conversionCache = new();

    /// <summary>
    ///     TTL for conversion cache entries.
    /// </summary>
    private readonly TimeSpan _conversionCacheTtl;

    /// <summary>
    ///     Document conversion service for format conversion.
    /// </summary>
    private readonly DocumentConversionService _conversionService;

    /// <summary>
    ///     Debounce delay for session modifications.
    /// </summary>
    private readonly TimeSpan _debounceDelay;

    /// <summary>
    ///     Cancellation token source for signaling disposal to async operations.
    /// </summary>
    private readonly CancellationTokenSource _disposeCts = new();

    /// <summary>
    ///     Extension manager for accessing extension instances.
    /// </summary>
    private readonly ExtensionManager _extensionManager;

    /// <summary>
    ///     Dictionary tracking last send times for frame skipping.
    /// </summary>
    private readonly ConcurrentDictionary<string, DateTime> _lastSendTimes = new();

    /// <summary>
    ///     Logger instance for diagnostic output.
    /// </summary>
    private readonly ILogger<ExtensionSessionBridge> _logger;

    /// <summary>
    ///     Maximum number of entries in the conversion cache to prevent memory exhaustion.
    /// </summary>
    private readonly int _maxConversionCacheSize;

    /// <summary>
    ///     Pending session modifications for debouncing.
    ///     Key: sessionId, Value: (requestor, timer).
    /// </summary>
    private readonly ConcurrentDictionary<string, (SessionIdentity Requestor, Timer Timer)> _pendingModifications =
        new();

    /// <summary>
    ///     Timer for periodic retry of pending snapshots.
    /// </summary>
    private readonly Timer _retryTimer;

    /// <summary>
    ///     Document session manager for accessing sessions.
    /// </summary>
    private readonly DocumentSessionManager _sessionManager;

    /// <summary>
    ///     Queue of stale locks that need to be disposed at shutdown.
    ///     Locks are queued here instead of immediate disposal to avoid race conditions.
    /// </summary>
    private readonly ConcurrentQueue<SemaphoreSlim> _staleLocks = new();

    /// <summary>
    ///     Counter for active timer callbacks to ensure graceful shutdown.
    /// </summary>
    private int _activeCallbacks;

    /// <summary>
    ///     Whether this instance has been disposed.
    /// </summary>
    private volatile bool _disposed;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtensionSessionBridge" /> class.
    /// </summary>
    /// <param name="config">Extension configuration.</param>
    /// <param name="sessionManager">Document session manager.</param>
    /// <param name="extensionManager">Extension manager.</param>
    /// <param name="conversionService">Document conversion service.</param>
    /// <param name="logger">Logger instance.</param>
    public ExtensionSessionBridge(
        ExtensionConfig config,
        DocumentSessionManager sessionManager,
        ExtensionManager extensionManager,
        DocumentConversionService conversionService,
        ILogger<ExtensionSessionBridge> logger)
    {
        _config = config;
        _sessionManager = sessionManager;
        _extensionManager = extensionManager;
        _conversionService = conversionService;
        _logger = logger;

        _debounceDelay = TimeSpan.FromMilliseconds(config.DebounceDelayMs);
        _conversionCacheTtl = TimeSpan.FromSeconds(config.ConversionCacheTtlSeconds);
        _maxConversionCacheSize = config.MaxConversionCacheSize;

        _sessionManager.SessionModified += HandleSessionModified;
        _sessionManager.SessionClosed += HandleSessionClosed;
        _extensionManager.ExtensionError += HandleExtensionError;

        _retryTimer = new Timer(
            OnRetryTimerElapsed,
            null,
            TimeSpan.FromSeconds(5),
            TimeSpan.FromSeconds(5));
    }

    /// <summary>
    ///     Disposes all resources and performs graceful shutdown.
    /// </summary>
    /// <remarks>
    ///     <para>Disposal sequence:</para>
    ///     <list type="number">
    ///         <item>Set _disposed flag to prevent new operations</item>
    ///         <item>Cancel the disposal token to signal async operations</item>
    ///         <item>Unsubscribe from session/extension events</item>
    ///         <item>Dispose the retry timer</item>
    ///         <item>Dispose pending modification timers</item>
    ///         <item>Wait for active callbacks and tasks (with timeout)</item>
    ///         <item>Dispose binding locks and stale locks</item>
    ///         <item>Clear all collections</item>
    ///     </list>
    /// </remarks>
    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;

        _disposeCts.Cancel();

        _sessionManager.SessionModified -= HandleSessionModified;
        _sessionManager.SessionClosed -= HandleSessionClosed;
        _extensionManager.ExtensionError -= HandleExtensionError;

        _retryTimer.Dispose();

        foreach (var pending in _pendingModifications.Values)
            pending.Timer.Dispose();
        _pendingModifications.Clear();

        WaitForActiveCallbacksAndTasks();

        foreach (var lockEntry in _bindingLocks)
            lockEntry.Value.Dispose();

        while (_staleLocks.TryDequeue(out var staleLock))
            staleLock.Dispose();

        _bindingLocks.Clear();
        _activeTasks.Clear();
        _conversionCache.Clear();
        _closedSessions.Clear();
        _disposeCts.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    ///     Called periodically to retry sending pending snapshots.
    /// </summary>
    /// <param name="state">Timer state (not used).</param>
    private void OnRetryTimerElapsed(object? state)
    {
        Interlocked.Increment(ref _activeCallbacks);
        try
        {
            if (_disposed)
                return;

            var taskId = Guid.NewGuid();
            CancellationToken cancellationToken;

            try
            {
                cancellationToken = _disposeCts.Token;
            }
            catch (ObjectDisposedException)
            {
                return;
            }

            if (_activeTasks.Count >= MaxActiveTasks)
            {
                _logger.LogWarning(
                    "Active task limit ({Limit}) reached, skipping retry timer",
                    MaxActiveTasks);
                return;
            }

            CleanupClosedSessionsSet();

            var task = Task.Run(async () =>
            {
                try
                {
                    await ProcessPendingSnapshotsAsync(cancellationToken);
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing pending snapshots");
                }
                finally
                {
                    _activeTasks.TryRemove(taskId, out _);
                }
            }, cancellationToken);
            _activeTasks.TryAdd(taskId, task);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unhandled exception in retry timer callback");
        }
        finally
        {
            Interlocked.Decrement(ref _activeCallbacks);
        }
    }

    /// <summary>
    ///     Waits for all active callbacks and tasks to complete during disposal.
    ///     Uses a hybrid approach with brief spins followed by sleep to reduce CPU usage.
    /// </summary>
    private void WaitForActiveCallbacksAndTasks()
    {
        var spinWait = new SpinWait();
        var timeout = DateTime.UtcNow.AddSeconds(1);
        var spinCount = 0;
        const int maxSpinsBeforeSleep = 10;

        while (Interlocked.CompareExchange(ref _activeCallbacks, 0, 0) > 0 &&
               DateTime.UtcNow < timeout)
            if (spinCount < maxSpinsBeforeSleep)
            {
                spinWait.SpinOnce();
                spinCount++;
            }
            else
            {
                Thread.Sleep(10);
            }

        var pendingTasks = _activeTasks.Values.ToArray();
        if (pendingTasks.Length > 0)
            try
            {
                Task.WaitAll(pendingTasks, TimeSpan.FromSeconds(5));
            }
            catch (AggregateException)
            {
            }
    }

    /// <summary>
    ///     Handles the SessionModified event from DocumentSessionManager.
    ///     Uses debouncing to prevent task flooding from rapid modifications.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="requestor">Requestor identity.</param>
    private void HandleSessionModified(string sessionId, SessionIdentity requestor)
    {
        if (_disposed || _bindings.Values.All(b => b.SessionId != sessionId))
            return;

        _pendingModifications.AddOrUpdate(
            sessionId,
            _ =>
            {
                var timer = new Timer(
                    OnDebounceTimerElapsed,
                    sessionId,
                    _debounceDelay,
                    Timeout.InfiniteTimeSpan);
                return (requestor, timer);
            },
            (_, existing) =>
            {
                try
                {
                    existing.Timer.Change(_debounceDelay, Timeout.InfiniteTimeSpan);
                }
                catch (ObjectDisposedException)
                {
                    var newTimer = new Timer(
                        OnDebounceTimerElapsed,
                        sessionId,
                        _debounceDelay,
                        Timeout.InfiniteTimeSpan);
                    return (requestor, newTimer);
                }

                return (requestor, existing.Timer);
            });
    }

    /// <summary>
    ///     Called when the debounce timer elapses for a session modification.
    /// </summary>
    /// <param name="state">Session ID.</param>
    private void OnDebounceTimerElapsed(object? state)
    {
        Interlocked.Increment(ref _activeCallbacks);
        try
        {
            if (_disposed || state is not string sessionId)
                return;

            if (!_pendingModifications.TryRemove(sessionId, out var pending))
                return;

            pending.Timer.Dispose();

            if (_activeTasks.Count >= MaxActiveTasks)
            {
                _logger.LogWarning(
                    "Active task limit ({Limit}) reached, dropping session modified event for {SessionId}",
                    MaxActiveTasks, sessionId);
                return;
            }

            var taskId = Guid.NewGuid();
            CancellationToken cancellationToken;

            try
            {
                cancellationToken = _disposeCts.Token;
            }
            catch (ObjectDisposedException)
            {
                return;
            }

            var task = Task.Run(async () =>
            {
                try
                {
                    await OnSessionModifiedAsync(sessionId, pending.Requestor, cancellationToken);
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    _logger.LogDebug(
                        "Session modified handling cancelled for session {SessionId}",
                        sessionId);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex,
                        "Error handling session modified event for session {SessionId}",
                        sessionId);
                }
                finally
                {
                    _activeTasks.TryRemove(taskId, out _);
                }
            }, cancellationToken);
            _activeTasks.TryAdd(taskId, task);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Unhandled exception in debounce timer callback");
        }
        finally
        {
            Interlocked.Decrement(ref _activeCallbacks);
        }
    }

    /// <summary>
    ///     Handles the ExtensionError event from ExtensionManager.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    private void HandleExtensionError(string extensionId)
    {
        if (_disposed)
            return;

        RemoveBindingsForExtension(extensionId);
    }

    /// <summary>
    ///     Handles the SessionClosed event from DocumentSessionManager.
    ///     Critical cleanup is performed synchronously to prevent orphan bindings.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="owner">Session owner identity.</param>
    private void HandleSessionClosed(string sessionId, SessionIdentity owner)
    {
        if (_disposed || _bindings.Values.All(b => b.SessionId != sessionId))
            return;

        _closedSessions[sessionId] = DateTime.UtcNow;
        CleanupPendingModification(sessionId);
        InvalidateSessionCache(sessionId);

        var bindings = _bindings.Values.Where(b => b.SessionId == sessionId).ToList();

        _ = UnbindAll(sessionId);

        if (bindings.Count == 0)
            return;

        if (_activeTasks.Count >= MaxActiveTasks)
        {
            _logger.LogWarning(
                "Active task limit ({Limit}) reached, skipping extension notifications for session {SessionId}. " +
                "Bindings were cleaned up synchronously.",
                MaxActiveTasks, sessionId);
            return;
        }

        var taskId = Guid.NewGuid();
        CancellationToken cancellationToken;

        try
        {
            cancellationToken = _disposeCts.Token;
        }
        catch (ObjectDisposedException)
        {
            return;
        }

        var task = Task.Run(async () =>
        {
            try
            {
                foreach (var binding in bindings)
                {
                    var extension = _extensionManager.GetRunningExtension(binding.ExtensionId);
                    if (extension != null)
                        await extension.NotifySessionClosedAsync(sessionId, ConvertOwner(owner), cancellationToken);
                }
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                _logger.LogDebug(
                    "Session closed notification cancelled for session {SessionId}",
                    sessionId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex,
                    "Error notifying extensions of session closed for session {SessionId}",
                    sessionId);
            }
            finally
            {
                _activeTasks.TryRemove(taskId, out _);
            }
        }, cancellationToken);
        _activeTasks.TryAdd(taskId, task);
    }

    /// <summary>
    ///     Binds a session to an extension with the specified output format.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="extensionId">Extension identifier.</param>
    /// <param name="outputFormat">Output format for conversion (e.g., "pdf", "html").</param>
    /// <param name="options">Conversion options for the binding.</param>
    /// <param name="requestor">Requestor identity for session access.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>Binding result.</returns>
    public async Task<BindingResult> BindAsync(
        string sessionId,
        string extensionId,
        string outputFormat,
        ConversionOptions options,
        SessionIdentity requestor,
        CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return BindingResult.Failure(ExtensionErrorCode.ExtensionDisabled, "Bridge has been disposed");

        if (_bindings.Count >= MaxBindings)
            return BindingResult.Failure(
                ExtensionErrorCode.InternalError,
                $"Maximum number of bindings ({MaxBindings}) reached. Unbind some sessions first.");

        var session = _sessionManager.TryGetSession(sessionId, requestor);
        if (session == null)
            return BindingResult.Failure(ExtensionErrorCode.SessionNotFound, $"Session not found: {sessionId}");

        var extension = await _extensionManager.GetExtensionAsync(extensionId);
        if (extension == null)
            return BindingResult.Failure(ExtensionErrorCode.ExtensionNotFound,
                $"Extension not found or unavailable: {extensionId}");

        if (extension.State == ExtensionState.Initializing)
            return BindingResult.Failure(
                ExtensionErrorCode.ExtensionInitializing,
                $"Extension '{extensionId}' is completing initialization. Please retry shortly.");

        if (!extension.CanHandle(session.Type.ToString().ToLowerInvariant(), outputFormat))
            return BindingResult.Failure(
                ExtensionErrorCode.FormatNotSupported,
                $"Extension '{extensionId}' cannot handle document type '{session.Type}' with format '{outputFormat}'");

        if (!_conversionService.IsFormatSupported(session.Type, outputFormat))
            return BindingResult.Failure(
                ExtensionErrorCode.FormatNotSupported,
                $"Format '{outputFormat}' is not supported for document type '{session.Type}'");

        if (_closedSessions.ContainsKey(sessionId))
            return BindingResult.Failure(
                ExtensionErrorCode.SessionNotFound,
                $"Session '{sessionId}' has been closed");

        var bindingKey = GetBindingKey(sessionId, extensionId);
        var binding = new SessionBindingInfo(_config.MaxConversionFailures, _config.FailureBackoffSeconds)
        {
            SessionId = sessionId,
            ExtensionId = extensionId,
            OutputFormat = outputFormat,
            ConversionOptions = options,
            Owner = session.Owner,
            CreatedAt = DateTime.UtcNow,
            LastSentAt = null
        };

        _bindings[bindingKey] = binding;

        if (_closedSessions.ContainsKey(sessionId))
        {
            _bindings.TryRemove(bindingKey, out _);
            return BindingResult.Failure(
                ExtensionErrorCode.SessionNotFound,
                $"Session '{sessionId}' was closed during binding creation");
        }

        _logger.LogInformation(
            "Bound session {SessionId} to extension {ExtensionId} with format {Format}",
            sessionId, extensionId, outputFormat);

        await SendSnapshotIfNeededAsync(session, extension, binding, cancellationToken);

        return BindingResult.Success(binding);
    }

    /// <summary>
    ///     Unbinds a session from an extension without notifying the extension.
    ///     Use <see cref="UnbindAndNotifyAsync" /> when notification is needed.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="extensionId">Extension identifier.</param>
    /// <returns>True if the binding was removed.</returns>
    public bool Unbind(string sessionId, string extensionId)
    {
        if (_disposed)
            return false;

        var bindingKey = GetBindingKey(sessionId, extensionId);
        var removed = _bindings.TryRemove(bindingKey, out _);

        if (removed)
        {
            CleanupBindingLock(bindingKey);
            _lastSendTimes.TryRemove(bindingKey, out _);
            _logger.LogInformation(
                "Unbound session {SessionId} from extension {ExtensionId}",
                sessionId, extensionId);
        }

        return removed;
    }

    /// <summary>
    ///     Unbinds a session from an extension and notifies the extension with session_unbound message.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="extensionId">Extension identifier.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>True if the binding was removed.</returns>
    public async Task<bool> UnbindAndNotifyAsync(
        string sessionId,
        string extensionId,
        CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return false;

        var bindingKey = GetBindingKey(sessionId, extensionId);
        var removed = _bindings.TryRemove(bindingKey, out var binding);

        if (removed)
        {
            CleanupBindingLock(bindingKey);
            _lastSendTimes.TryRemove(bindingKey, out _);
            _logger.LogInformation(
                "Unbound session {SessionId} from extension {ExtensionId}",
                sessionId, extensionId);

            var extension = await _extensionManager.GetExtensionAsync(extensionId);
            if (extension != null)
                await extension.NotifySessionUnboundAsync(
                    sessionId,
                    binding != null ? ConvertOwner(binding.Owner) : null,
                    cancellationToken);
        }

        return removed;
    }

    /// <summary>
    ///     Unbinds all extensions from a session without notifying extensions.
    ///     Use <see cref="UnbindAllAndNotifyAsync" /> when notification is needed.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <returns>Number of bindings removed.</returns>
    public int UnbindAll(string sessionId)
    {
        if (_disposed)
            return 0;

        var keysToRemove = _bindings.Keys
            .Where(k => k.StartsWith($"{sessionId}{BindingKeySeparator}", StringComparison.Ordinal))
            .ToList();

        foreach (var key in keysToRemove)
        {
            _bindings.TryRemove(key, out _);
            CleanupBindingLock(key);
            _lastSendTimes.TryRemove(key, out _);
        }

        if (keysToRemove.Count > 0)
            _logger.LogInformation(
                "Unbound all {Count} extension(s) from session {SessionId}",
                keysToRemove.Count, sessionId);

        return keysToRemove.Count;
    }

    /// <summary>
    ///     Unbinds all extensions from a session and notifies each extension with session_unbound message.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>Number of bindings removed.</returns>
    public async Task<int> UnbindAllAndNotifyAsync(string sessionId, CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return 0;

        var keysToRemove = _bindings.Keys
            .Where(k => k.StartsWith($"{sessionId}{BindingKeySeparator}", StringComparison.Ordinal))
            .ToList();

        var bindingsToNotify = new List<(string ExtensionId, SessionOwner? Owner)>();

        foreach (var key in keysToRemove)
            if (_bindings.TryRemove(key, out var binding))
            {
                CleanupBindingLock(key);
                _lastSendTimes.TryRemove(key, out _);
                bindingsToNotify.Add((binding.ExtensionId, ConvertOwner(binding.Owner)));
            }

        if (bindingsToNotify.Count > 0)
        {
            _logger.LogInformation(
                "Unbound all {Count} extension(s) from session {SessionId}",
                bindingsToNotify.Count, sessionId);

            foreach (var (extensionId, owner) in bindingsToNotify)
            {
                var extension = await _extensionManager.GetExtensionAsync(extensionId);
                if (extension != null)
                    await extension.NotifySessionUnboundAsync(sessionId, owner, cancellationToken);
            }
        }

        return bindingsToNotify.Count;
    }

    /// <summary>
    ///     Changes the output format for an existing binding.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="extensionId">Extension identifier.</param>
    /// <param name="newFormat">New output format.</param>
    /// <param name="options">Conversion options for the binding.</param>
    /// <param name="requestor">Requestor identity.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>Result of the format change.</returns>
    public async Task<BindingResult> SetFormatAsync(
        string sessionId,
        string extensionId,
        string newFormat,
        ConversionOptions options,
        SessionIdentity requestor,
        CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return BindingResult.Failure(ExtensionErrorCode.ExtensionDisabled, "Bridge has been disposed");

        var bindingKey = GetBindingKey(sessionId, extensionId);

        if (!_bindings.TryGetValue(bindingKey, out var binding))
            return BindingResult.Failure(
                ExtensionErrorCode.BindingNotFound,
                $"No binding found for session '{sessionId}' and extension '{extensionId}'");

        var session = _sessionManager.TryGetSession(sessionId, requestor);
        if (session == null)
            return BindingResult.Failure(ExtensionErrorCode.SessionNotFound, $"Session not found: {sessionId}");

        if (!_conversionService.IsFormatSupported(session.Type, newFormat))
            return BindingResult.Failure(
                ExtensionErrorCode.FormatNotSupported,
                $"Format '{newFormat}' is not supported for document type '{session.Type}'");

        var extension = await _extensionManager.GetExtensionAsync(extensionId);
        if (extension == null)
            return BindingResult.Failure(ExtensionErrorCode.ExtensionNotFound,
                $"Extension not found or unavailable: {extensionId}");

        if (!extension.CanHandle(session.Type.ToString().ToLowerInvariant(), newFormat))
            return BindingResult.Failure(
                ExtensionErrorCode.FormatNotSupported,
                $"Extension '{extensionId}' cannot handle format '{newFormat}'");

        var bindingLock = GetOrCreateBindingLock(bindingKey);
        await bindingLock.WaitAsync(cancellationToken);
        try
        {
            if (!_bindings.ContainsKey(bindingKey))
                return BindingResult.Failure(
                    ExtensionErrorCode.BindingNotFound,
                    "Binding was removed during format change");

            binding.UpdateFormatAndOptions(newFormat, options);

            _logger.LogInformation(
                "Changed format for session {SessionId} / extension {ExtensionId} to {Format}",
                sessionId, extensionId, newFormat);
        }
        finally
        {
            bindingLock.Release();
        }

        await SendSnapshotIfNeededAsync(session, extension, binding, cancellationToken);

        return BindingResult.Success(binding);
    }

    /// <summary>
    ///     Gets all bindings for a session.
    ///     Returns a snapshot copy to prevent external modification.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <returns>List of bindings.</returns>
    public IReadOnlyList<SessionBindingInfo> GetBindings(string sessionId)
    {
        if (_disposed)
            return [];

        return _bindings.Values.Where(b => b.SessionId == sessionId).ToList();
    }

    /// <summary>
    ///     Gets all bindings for an extension.
    ///     Returns a snapshot copy to prevent external modification.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    /// <returns>List of bindings.</returns>
    public IReadOnlyList<SessionBindingInfo> GetBindingsByExtension(string extensionId)
    {
        if (_disposed)
            return [];

        return _bindings.Values.Where(b => b.ExtensionId == extensionId).ToList();
    }

    /// <summary>
    ///     Removes all bindings for an extension.
    ///     Called when an extension enters Error state to clean up resources.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    /// <returns>Number of bindings removed.</returns>
    // ReSharper disable once UnusedMethodReturnValue.Global - Public API, return value useful for callers
    public int RemoveBindingsForExtension(string extensionId)
    {
        if (_disposed)
            return 0;

        var keysToRemove = _bindings.Keys
            .Where(k => k.EndsWith($"{BindingKeySeparator}{extensionId}", StringComparison.Ordinal))
            .ToList();

        foreach (var key in keysToRemove)
        {
            _bindings.TryRemove(key, out _);
            CleanupBindingLock(key);
            _lastSendTimes.TryRemove(key, out _);
        }

        if (keysToRemove.Count > 0)
            _logger.LogInformation(
                "Removed {Count} binding(s) for extension {ExtensionId} due to extension error",
                keysToRemove.Count, extensionId);

        return keysToRemove.Count;
    }

    /// <summary>
    ///     Notifies the bridge that a session has been modified.
    ///     Call this when IsDirty becomes true.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="requestor">Requestor identity.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public async Task OnSessionModifiedAsync(
        string sessionId,
        SessionIdentity requestor,
        CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return;

        var session = _sessionManager.TryGetSession(sessionId, requestor);
        if (session == null)
            return;

        InvalidateSessionCache(sessionId);

        var bindings = _bindings.Values.Where(b => b.SessionId == sessionId).ToList();

        foreach (var binding in bindings)
        {
            binding.NeedsSend = true;
            await TrySendPendingSnapshotAsync(session, binding, cancellationToken);
        }
    }

    /// <summary>
    ///     Invalidates all cached conversion results for a session.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    private void InvalidateSessionCache(string sessionId)
    {
        var keysToRemove = _conversionCache.Keys
            .Where(k => k.StartsWith($"{sessionId}{BindingKeySeparator}", StringComparison.Ordinal))
            .ToList();

        foreach (var key in keysToRemove)
            _conversionCache.TryRemove(key, out _);

        if (keysToRemove.Count > 0)
            _logger.LogDebug(
                "Invalidated {Count} cached conversion(s) for session {SessionId}",
                keysToRemove.Count, sessionId);
    }

    /// <summary>
    ///     Attempts to send a pending snapshot for a binding if the extension is available.
    /// </summary>
    /// <param name="session">The document session.</param>
    /// <param name="binding">The binding information.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    private async Task TrySendPendingSnapshotAsync(
        DocumentSession session,
        SessionBindingInfo binding,
        CancellationToken cancellationToken)
    {
        if (_disposed || !binding.NeedsSend)
            return;

        if (binding.IsInBackoff())
        {
            _logger.LogDebug(
                "Binding {SessionId}/{ExtensionId} is in backoff due to conversion failures",
                binding.SessionId, binding.ExtensionId);
            return;
        }

        var extension = await _extensionManager.GetExtensionAsync(binding.ExtensionId);
        if (extension == null)
            return;

        if (extension.State == ExtensionState.Busy)
        {
            _logger.LogDebug(
                "Extension {ExtensionId} is busy, deferring snapshot for session {SessionId}",
                binding.ExtensionId, binding.SessionId);
            return;
        }

        await SendSnapshotIfNeededAsync(session, extension, binding, cancellationToken);
    }

    /// <summary>
    ///     Processes all bindings with pending snapshots (NeedsSend=true).
    ///     Called to retry deferred snapshots when extensions become available.
    ///     Uses a timeout to prevent unbounded processing when many bindings are pending.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token.</param>
    public async Task ProcessPendingSnapshotsAsync(CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return;

        var pendingBindings = _bindings.Values.Where(b => b.NeedsSend).ToList();
        if (pendingBindings.Count == 0)
            return;

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(TimeSpan.FromSeconds(_config.RetryLoopTimeoutSeconds));

        var processedCount = 0;

        foreach (var binding in pendingBindings)
        {
            if (timeoutCts.Token.IsCancellationRequested)
            {
                if (!cancellationToken.IsCancellationRequested)
                    _logger.LogDebug(
                        "Retry loop timeout reached after processing {Processed}/{Total} pending bindings",
                        processedCount, pendingBindings.Count);
                break;
            }

            var session = _sessionManager.TryGetSession(binding.SessionId, binding.Owner);
            if (session == null)
            {
                _logger.LogDebug(
                    "Session {SessionId} no longer exists, notifying and removing orphan binding to {ExtensionId}",
                    binding.SessionId, binding.ExtensionId);

                var extension = _extensionManager.GetRunningExtension(binding.ExtensionId);
                if (extension != null)
                    await extension.NotifySessionClosedAsync(
                        binding.SessionId,
                        ConvertOwner(binding.Owner),
                        timeoutCts.Token);

                _ = Unbind(binding.SessionId, binding.ExtensionId);
                continue;
            }

            await TrySendPendingSnapshotAsync(session, binding, timeoutCts.Token);
            processedCount++;
        }
    }

    /// <summary>
    ///     Notifies extensions that a session has been closed and performs cleanup.
    ///     This method is kept for API compatibility.
    /// </summary>
    /// <remarks>
    ///     If the session was already closed via the SessionClosed event (HandleSessionClosed),
    ///     this method will only notify extensions that weren't notified yet.
    ///     Critical cleanup (unbinding) is done synchronously to prevent orphan bindings.
    /// </remarks>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="owner">Session owner identity.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public async Task OnSessionClosedAsync(
        string sessionId,
        SessionIdentity owner,
        CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return;

        var alreadyProcessed = _closedSessions.ContainsKey(sessionId);

        if (!alreadyProcessed)
        {
            _closedSessions[sessionId] = DateTime.UtcNow;
            CleanupPendingModification(sessionId);
            InvalidateSessionCache(sessionId);
        }

        var bindings = _bindings.Values.Where(b => b.SessionId == sessionId).ToList();

        _ = UnbindAll(sessionId);

        foreach (var binding in bindings)
        {
            var extension = _extensionManager.GetRunningExtension(binding.ExtensionId);
            if (extension != null)
                await extension.NotifySessionClosedAsync(sessionId, ConvertOwner(owner), cancellationToken);
        }
    }

    /// <summary>
    ///     Cleans up any pending debounce modification for a session.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    private void CleanupPendingModification(string sessionId)
    {
        if (_pendingModifications.TryRemove(sessionId, out var pending))
        {
            pending.Timer.Dispose();
            _logger.LogDebug(
                "Cleaned up pending modification for closed session {SessionId}",
                sessionId);
        }
    }

    /// <summary>
    ///     Cleans up old entries from the closed sessions set to prevent unbounded growth.
    ///     Entries older than the conversion cache TTL are safe to remove.
    /// </summary>
    private void CleanupClosedSessionsSet()
    {
        var threshold = DateTime.UtcNow - _conversionCacheTtl - TimeSpan.FromSeconds(10);
        var keysToRemove = _closedSessions
            .Where(kvp => kvp.Value < threshold)
            .Select(kvp => kvp.Key)
            .ToList();

        foreach (var key in keysToRemove)
            _closedSessions.TryRemove(key, out _);
    }

    /// <summary>
    ///     Sends a snapshot to the extension if the frame interval has elapsed.
    ///     Implements frame skipping to avoid overwhelming extensions with rapid updates.
    /// </summary>
    /// <param name="session">The document session to snapshot.</param>
    /// <param name="extension">The target extension.</param>
    /// <param name="binding">The binding information.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>A task representing the operation.</returns>
    private async Task SendSnapshotIfNeededAsync(
        DocumentSession session,
        Extension extension,
        SessionBindingInfo binding,
        CancellationToken cancellationToken)
    {
        var bindingKey = GetBindingKey(binding.SessionId, binding.ExtensionId);
        var now = DateTime.UtcNow;

        var effectiveMinFrameIntervalMs = extension.Definition.GetEffectiveFrameIntervalMs(_config);

        if (_lastSendTimes.TryGetValue(bindingKey, out var lastSend))
        {
            var elapsed = now - lastSend;
            if (elapsed.TotalMilliseconds < effectiveMinFrameIntervalMs)
            {
                _logger.LogDebug(
                    "Skipping snapshot for {SessionId}/{ExtensionId}, frame interval not elapsed",
                    binding.SessionId, binding.ExtensionId);
                return;
            }
        }

        var outputFormat = binding.OutputFormat;
        var conversionOptions = binding.ConversionOptions;
        var optionsCacheKey = binding.GetOptionsCacheKey();
        var data = await ConvertSessionAsync(session, outputFormat, conversionOptions, optionsCacheKey,
            cancellationToken);
        if (data == null)
        {
            var inBackoff = binding.RecordConversionFailure();
            if (inBackoff)
                _logger.LogWarning(
                    "Conversion for session {SessionId} to format {Format} failed {Count} times, " +
                    "entering backoff for {BackoffSeconds} seconds",
                    binding.SessionId, outputFormat, binding.ConversionFailures, _config.FailureBackoffSeconds);
            else
                _logger.LogWarning(
                    "Failed to convert session {SessionId} to format {Format} (attempt {Count})",
                    binding.SessionId, outputFormat, binding.ConversionFailures);
            return;
        }

        var bindingLock = GetOrCreateBindingLock(bindingKey);

        await bindingLock.WaitAsync(cancellationToken);
        try
        {
            if (!_bindings.ContainsKey(bindingKey))
            {
                _logger.LogDebug(
                    "Binding was removed while waiting for lock: {BindingKey}",
                    bindingKey);
                return;
            }

            if (extension.State is ExtensionState.Error or ExtensionState.Stopping or ExtensionState.Unloaded)
            {
                _logger.LogDebug(
                    "Extension {ExtensionId} is no longer available after lock acquisition, state: {State}",
                    binding.ExtensionId, extension.State);
                return;
            }

            now = DateTime.UtcNow;
            if (_lastSendTimes.TryGetValue(bindingKey, out lastSend))
            {
                var elapsed = now - lastSend;
                if (elapsed.TotalMilliseconds < effectiveMinFrameIntervalMs)
                    return;
            }

            var currentFormat = binding.OutputFormat;
            if (currentFormat != outputFormat)
            {
                _logger.LogDebug(
                    "Format changed from {OldFormat} to {NewFormat} during conversion for {SessionId}/{ExtensionId}, skipping",
                    outputFormat, currentFormat, binding.SessionId, binding.ExtensionId);
                return;
            }

            var metadata = CreateMetadata(session, outputFormat);
            var success = await extension.SendSnapshotAsync(data, metadata, cancellationToken);

            if (success)
            {
                binding.UpdateLastSent(now);
                _lastSendTimes[bindingKey] = now;
            }
        }
        finally
        {
            bindingLock.Release();
        }
    }

    /// <summary>
    ///     Gets or creates a semaphore lock for a specific binding.
    /// </summary>
    /// <param name="bindingKey">The binding key.</param>
    /// <returns>A semaphore for the binding.</returns>
    private SemaphoreSlim GetOrCreateBindingLock(string bindingKey)
    {
        return _bindingLocks.GetOrAdd(bindingKey, _ => new SemaphoreSlim(1, 1));
    }

    /// <summary>
    ///     Cleans up the lock for a removed binding.
    ///     Does not dispose the lock immediately to avoid race conditions with active waiters.
    ///     Locks are collected in _staleLocks and disposed at shutdown.
    ///     Periodically cleans up excess stale locks to prevent unbounded growth.
    /// </summary>
    /// <param name="bindingKey">The binding key.</param>
    private void CleanupBindingLock(string bindingKey)
    {
        if (_bindingLocks.TryRemove(bindingKey, out var staleLock))
            _staleLocks.Enqueue(staleLock);

        if (_staleLocks.Count > MaxStaleLocks)
            DrainStaleLocks(MaxStaleLocks / 2);
    }

    /// <summary>
    ///     Drains and disposes stale locks until the count is at or below the target.
    /// </summary>
    /// <param name="targetCount">Target count to drain to.</param>
    private void DrainStaleLocks(int targetCount)
    {
        while (_staleLocks.Count > targetCount && _staleLocks.TryDequeue(out var lockToDispose))
            try
            {
                lockToDispose.Dispose();
            }
            // ReSharper disable once EmptyGeneralCatchClause - Best-effort disposal during cleanup
            catch
            {
            }
    }

    /// <summary>
    ///     Converts a document session to the specified output format.
    /// </summary>
    /// <param name="session">The document session to convert.</param>
    /// <param name="outputFormat">The target output format.</param>
    /// <param name="options">Conversion options.</param>
    /// <param name="optionsCacheKey">Options hash for cache key.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>
    ///     The converted document as a byte array, or <c>null</c> if conversion failed.
    /// </returns>
    private async Task<byte[]?> ConvertSessionAsync(
        DocumentSession session,
        string outputFormat,
        ConversionOptions options,
        int optionsCacheKey,
        CancellationToken cancellationToken)
    {
        var cacheKey = $"{session.SessionId}{BindingKeySeparator}{outputFormat}{BindingKeySeparator}{optionsCacheKey}";
        var now = DateTime.UtcNow;

        if (_conversionCache.TryGetValue(cacheKey, out var cached) &&
            now - cached.Timestamp < _conversionCacheTtl)
        {
            _conversionCache[cacheKey] = (cached.Data, now);
            _logger.LogDebug(
                "Using cached conversion for session {SessionId} format {Format}",
                session.SessionId, outputFormat);
            return cached.Data;
        }

        try
        {
            var data = await session.ExecuteAsync(
                doc => _conversionService.ConvertToBytes(doc, session.Type, outputFormat, options),
                cancellationToken);

            if (!_closedSessions.ContainsKey(session.SessionId))
            {
                EnsureCacheCapacity();
                _conversionCache[cacheKey] = (data, now);
            }

            return data;
        }
        catch (ObjectDisposedException)
        {
            _logger.LogDebug(
                "Session {SessionId} was closed during conversion to {Format}",
                session.SessionId, outputFormat);
            return null;
        }
        catch (OperationCanceledException)
        {
            _logger.LogDebug(
                "Conversion of session {SessionId} to {Format} was cancelled",
                session.SessionId, outputFormat);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex,
                "Failed to convert session {SessionId} to {Format}",
                session.SessionId, outputFormat);
            return null;
        }
    }

    /// <summary>
    ///     Ensures the conversion cache has capacity for a new entry by removing
    ///     expired entries and evicting oldest entries if at size limit.
    ///     Called before adding to prevent cache from exceeding size limit.
    ///     Uses fresh cache count after each operation to handle concurrent modifications.
    /// </summary>
    private void EnsureCacheCapacity()
    {
        var now = DateTime.UtcNow;
        var cacheSnapshot = _conversionCache.ToList();
        var expiredKeys = cacheSnapshot
            .Where(kvp => now - kvp.Value.Timestamp > _conversionCacheTtl)
            .Select(kvp => kvp.Key)
            .ToList();

        foreach (var key in expiredKeys)
            _conversionCache.TryRemove(key, out _);

        var currentCount = _conversionCache.Count;
        if (currentCount >= _maxConversionCacheSize)
        {
            var freshSnapshot = _conversionCache.ToList();
            var excessCount = freshSnapshot.Count - _maxConversionCacheSize + 1;

            if (excessCount > 0)
            {
                var oldestKeys = freshSnapshot
                    .OrderBy(kvp => kvp.Value.Timestamp)
                    .Take(excessCount)
                    .Select(kvp => kvp.Key)
                    .ToList();

                foreach (var key in oldestKeys)
                    _conversionCache.TryRemove(key, out _);

                if (oldestKeys.Count > 0)
                    _logger.LogDebug(
                        "Evicted {Count} oldest cache entries to ensure capacity",
                        oldestKeys.Count);
            }
        }

        var hardLimit = _maxConversionCacheSize + _maxConversionCacheSize / 10;
        if (_conversionCache.Count > hardLimit)
        {
            var keysToRemove = _conversionCache
                .OrderBy(kvp => kvp.Value.Timestamp)
                .Take(_conversionCache.Count - _maxConversionCacheSize)
                .Select(kvp => kvp.Key)
                .ToList();

            foreach (var key in keysToRemove)
                _conversionCache.TryRemove(key, out _);

            if (keysToRemove.Count > 0)
                _logger.LogDebug(
                    "Hard limit cleanup: evicted {Count} cache entries",
                    keysToRemove.Count);
        }
    }

    /// <summary>
    ///     Creates metadata for a snapshot from session information and output format.
    /// </summary>
    /// <param name="session">The document session.</param>
    /// <param name="outputFormat">The output format captured at the time of conversion.</param>
    /// <returns>A new <see cref="ExtensionMetadata" /> instance.</returns>
    private ExtensionMetadata CreateMetadata(DocumentSession session, string outputFormat)
    {
        return new ExtensionMetadata
        {
            SessionId = session.SessionId,
            DocumentType = session.Type.ToString().ToLowerInvariant(),
            OriginalPath = session.Path,
            OutputFormat = outputFormat,
            MimeType = _conversionService.GetMimeType(outputFormat),
            Timestamp = DateTime.UtcNow,
            Owner = ConvertOwner(session.Owner)
        };
    }

    /// <summary>
    ///     Converts a <see cref="SessionIdentity" /> to a <see cref="SessionOwner" />.
    /// </summary>
    /// <param name="identity">The session identity to convert.</param>
    /// <returns>
    ///     A <see cref="SessionOwner" /> instance, or <c>null</c> if the identity is anonymous.
    /// </returns>
    private static SessionOwner? ConvertOwner(SessionIdentity identity)
    {
        if (identity.IsAnonymous)
            return null;

        return new SessionOwner
        {
            GroupId = identity.GroupId,
            UserId = identity.UserId
        };
    }

    /// <summary>
    ///     Generates a unique key for a session-extension binding.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <param name="extensionId">The extension identifier.</param>
    /// <returns>A composite key string in the format "sessionId:extensionId".</returns>
    private static string GetBindingKey(string sessionId, string extensionId)
    {
        return $"{sessionId}{BindingKeySeparator}{extensionId}";
    }
}
