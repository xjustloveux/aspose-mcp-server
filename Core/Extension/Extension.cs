using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using AsposeMcpServer.Core.Extension.Transport;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Represents a single extension instance with process lifecycle management.
///     Handles process startup, shutdown, heartbeat, snapshot sending, and restart logic.
/// </summary>
public class Extension : IAsyncDisposable
{
    /// <summary>
    ///     Threshold in seconds - if crash happens within this time of start, it's a rapid crash.
    /// </summary>
    /// <remarks>
    ///     10 seconds was chosen as a balance between:
    ///     - Being short enough to detect startup failures (config errors, missing dependencies)
    ///     - Being long enough to exclude normal operation crashes
    ///     This value is intentionally not configurable to provide consistent crash detection behavior.
    /// </remarks>
    private const int RapidCrashThresholdSeconds = 10;

    /// <summary>
    ///     Maximum consecutive rapid crashes before entering Error state.
    /// </summary>
    /// <remarks>
    ///     3 attempts allows for transient issues (e.g., temporary port conflicts, file locks)
    ///     while preventing infinite restart loops for persistent configuration problems.
    ///     Combined with RapidCrashThresholdSeconds, this means 3 crashes within ~30 seconds
    ///     of cumulative runtime will trigger Error state.
    /// </remarks>
    private const int MaxRapidCrashes = 3;

    /// <summary>
    ///     Timeout in seconds for waiting on reader tasks during cleanup.
    ///     Longer timeout after process kill to ensure streams are properly drained.
    /// </summary>
    private const int ReaderTaskCleanupTimeoutSeconds = 10;

    /// <summary>
    ///     Extension configuration containing restart limits and timeouts.
    /// </summary>
    private readonly ExtensionConfig _config;

    /// <summary>
    ///     The extension definition from extensions.json.
    /// </summary>
    private readonly ExtensionDefinition _definition;

    /// <summary>
    ///     Logger instance for diagnostic output.
    /// </summary>
    private readonly ILogger _logger;

    /// <summary>
    ///     Pending commands awaiting responses, keyed by command ID.
    /// </summary>
    private readonly ConcurrentDictionary<string, TaskCompletionSource<CommandResponse>> _pendingCommands = new();

    /// <summary>
    ///     Process cleanup manager for registering this extension's process.
    /// </summary>
    private readonly ProcessCleanupManager? _processCleanupManager;

    /// <summary>
    ///     Semaphore to synchronize process lifecycle operations.
    /// </summary>
    private readonly SemaphoreSlim _processLock = new(1, 1);

    /// <summary>
    ///     Snapshot manager for tracking pending snapshots.
    /// </summary>
    private readonly SnapshotManager _snapshotManager;

    /// <summary>
    ///     Lock object for state transitions.
    /// </summary>
    private readonly object _stateLock = new();

    /// <summary>
    ///     Semaphore to synchronize StandardInput write operations.
    ///     Prevents message interleaving when multiple operations write concurrently.
    /// </summary>
    private readonly SemaphoreSlim _stdinLock = new(1, 1);

    /// <summary>
    ///     Transport for sending data to the extension process.
    /// </summary>
    private readonly IExtensionTransport _transport;

    /// <summary>
    ///     Counter for active snapshot operations. Used for proper Busy/Idle state management.
    /// </summary>
    private int _activeOperations;

    /// <summary>
    ///     Counter for consecutive send failures. Thread-safe via Interlocked.
    /// </summary>
    private int _consecutiveFailures;

    /// <summary>
    ///     Whether this instance has been disposed.
    /// </summary>
    private volatile bool _disposed;

    /// <summary>
    ///     Flag indicating a restart is in progress to prevent concurrent restarts.
    /// </summary>
    private int _isRestarting;

    /// <summary>
    ///     Timestamp of the last activity (send or receive) stored as ticks for thread-safe access.
    /// </summary>
    private long _lastActivityTicks = DateTime.UtcNow.Ticks;

    /// <summary>
    ///     Timestamp of the last heartbeat response received, stored as ticks for thread-safe access.
    ///     This is separate from _lastActivityTicks to accurately track heartbeat responsiveness
    ///     even when there is other activity (e.g., snapshot sends).
    /// </summary>
    /// <remarks>
    ///     Initialized to process start time to provide a grace period for the first heartbeat.
    ///     This prevents false "missed heartbeat" detection immediately after startup.
    ///     Updated when a "pong" message is received from the extension.
    /// </remarks>
    private long _lastHeartbeatResponseTicks;

    /// <summary>
    ///     Timestamp of the last heartbeat sent, stored as ticks for thread-safe access.
    ///     Used for detecting heartbeat timeout.
    /// </summary>
    /// <remarks>
    ///     Initialized to 0 to indicate no heartbeat has been sent yet.
    ///     Only updated after a heartbeat message is successfully written to stdin.
    ///     Compare with <see cref="_lastHeartbeatResponseTicks" /> to detect missed heartbeats.
    /// </remarks>
    private long _lastHeartbeatSentTicks;

    /// <summary>
    ///     Number of consecutive missed heartbeat responses.
    /// </summary>
    /// <remarks>
    ///     A heartbeat is considered "missed" if <see cref="_lastHeartbeatResponseTicks" /> is
    ///     less than <see cref="_lastHeartbeatSentTicks" /> when sending the next heartbeat.
    ///     Reset to 0 when a pong is received. When this reaches <see cref="ExtensionConfig.MaxMissedHeartbeats" />,
    ///     the extension is marked as crashed.
    /// </remarks>
    private int _missedHeartbeats;

    /// <summary>
    ///     The extension process instance.
    /// </summary>
    private Process? _process;

    /// <summary>
    ///     Generation counter for detecting process replacement.
    ///     Incremented each time a new process is started.
    /// </summary>
    private long _processGeneration;

    /// <summary>
    ///     Timestamp when the current process was started, stored as ticks.
    /// </summary>
    private long _processStartTimeTicks;

    /// <summary>
    ///     Counter for consecutive rapid crashes (crashes within RapidCrashThresholdSeconds of start).
    /// </summary>
    private int _rapidCrashCount;

    /// <summary>
    ///     Cancellation token source for reader tasks.
    ///     Used to signal graceful shutdown of reader loops.
    /// </summary>
    private CancellationTokenSource? _readerCts;

    /// <summary>
    ///     Counter for restart attempts.
    /// </summary>
    private int _restartCount;

    /// <summary>
    ///     Sequence number for snapshot tracking.
    /// </summary>
    private long _sequenceNumber;

    /// <summary>
    ///     Current state of the extension. Volatile for thread-safe reads.
    /// </summary>
    private volatile ExtensionState _state = ExtensionState.Unloaded;

    /// <summary>
    ///     Background task reading stderr from the extension process.
    /// </summary>
    private Task? _stderrReaderTask;

    /// <summary>
    ///     Background task reading stdout from the extension process.
    /// </summary>
    private Task? _stdoutReaderTask;

    /// <summary>
    ///     Initializes a new instance of the <see cref="Extension" /> class.
    /// </summary>
    /// <param name="definition">Extension definition.</param>
    /// <param name="config">Extension configuration.</param>
    /// <param name="transport">Transport for communication.</param>
    /// <param name="snapshotManager">Snapshot manager for resource tracking.</param>
    /// <param name="processCleanupManager">Process cleanup manager for registering processes.</param>
    /// <param name="logger">Logger instance.</param>
    public Extension(
        ExtensionDefinition definition,
        ExtensionConfig config,
        IExtensionTransport transport,
        SnapshotManager snapshotManager,
        ProcessCleanupManager? processCleanupManager,
        ILogger logger)
    {
        _definition = definition;
        _config = config;
        _transport = transport;
        _snapshotManager = snapshotManager;
        _processCleanupManager = processCleanupManager;
        _logger = logger;
    }

    /// <summary>
    ///     Gets the extension definition.
    /// </summary>
    public ExtensionDefinition Definition => _definition;

    /// <summary>
    ///     Gets the effective max missed heartbeats for this extension.
    ///     Uses extension-specific setting if configured, constrained by global limits.
    /// </summary>
    private int EffectiveMaxMissedHeartbeats => _definition.GetEffectiveMaxMissedHeartbeats(_config);

    /// <summary>
    ///     Gets the effective snapshot TTL in seconds for this extension.
    ///     Uses extension-specific setting if configured, constrained by global limits.
    /// </summary>
    private int EffectiveSnapshotTtlSeconds => _definition.GetEffectiveSnapshotTtlSeconds(_config);

    /// <summary>
    ///     Gets the current state of the extension.
    /// </summary>
    public ExtensionState State => _state;

    /// <summary>
    ///     Gets the time of last activity.
    /// </summary>
    public DateTime LastActivity => new(Interlocked.Read(ref _lastActivityTicks), DateTimeKind.Utc);

    /// <summary>
    ///     Gets the number of restart attempts. Thread-safe.
    /// </summary>
    public int RestartCount => Interlocked.CompareExchange(ref _restartCount, 0, 0);

    /// <inheritdoc />
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        _disposed = true;

        try
        {
            await StopAsync();
        }
        finally
        {
            MessageReceived = null;
            StateChanged = null;

            if (_readerCts != null)
            {
                try
                {
                    await _readerCts.CancelAsync();
                    _readerCts.Dispose();
                }
                catch
                {
                    // Ignore disposal errors during cleanup
                }

                _readerCts = null;
            }

            _processLock.Dispose();
            _stdinLock.Dispose();
        }

        GC.SuppressFinalize(this);
    }

    /// <summary>
    ///     Event raised when a message is received from the extension.
    /// </summary>
    public event EventHandler<ExtensionMessageEventArgs>? MessageReceived;

    /// <summary>
    ///     Event raised when the extension state changes.
    /// </summary>
    public event EventHandler<ExtensionState>? StateChanged;

    /// <summary>
    ///     Checks if this extension can handle the given document type and output format.
    /// </summary>
    /// <param name="documentType">Document type (e.g., "word", "excel").</param>
    /// <param name="outputFormat">Output format (e.g., "pdf", "html").</param>
    /// <returns>True if the extension can handle the combination.</returns>
    public bool CanHandle(string documentType, string outputFormat)
    {
        if (!_definition.IsAvailable)
            return false;

        var supportsDocType = _definition.SupportedDocumentTypes.Count == 0 ||
                              _definition.SupportedDocumentTypes.Contains(documentType,
                                  StringComparer.OrdinalIgnoreCase);

        var supportsFormat = _definition.InputFormats.Count == 0 ||
                             _definition.InputFormats.Contains(outputFormat, StringComparer.OrdinalIgnoreCase);

        return supportsDocType && supportsFormat;
    }

    /// <summary>
    ///     Ensures the extension process is started (lazy loading).
    /// </summary>
    /// <returns>True if the extension is running; otherwise, false.</returns>
    public async Task<bool> EnsureStartedAsync()
    {
        if (_disposed)
            return false;

        if (_state == ExtensionState.Idle || _state == ExtensionState.Busy)
            return true;

        if (_state == ExtensionState.Initializing)
            return await WaitForInitializationAsync();

        if (_state == ExtensionState.Error)
        {
            _logger.LogWarning(
                "Extension {ExtensionId} is in Error state and cannot be started. " +
                "Restart the server or reload extension configuration to recover.",
                _definition.Id);
            return false;
        }

        if (_state == ExtensionState.Crashed)
        {
            _logger.LogDebug(
                "Extension {ExtensionId} is in Crashed state, waiting for restart handler",
                _definition.Id);
            return false;
        }

        if (_state == ExtensionState.Stopping)
        {
            _logger.LogDebug(
                "Extension {ExtensionId} is stopping, cannot start now",
                _definition.Id);
            return false;
        }

        try
        {
            await _processLock.WaitAsync();
        }
        catch (ObjectDisposedException)
        {
            return false;
        }

        try
        {
            switch (_state)
            {
                case ExtensionState.Idle:
                case ExtensionState.Busy:
                    return true;

                case ExtensionState.Initializing:
                    return await WaitForInitializationAsync();

                case ExtensionState.Error:
                case ExtensionState.Crashed:
                case ExtensionState.Stopping:
                    return false;

                case ExtensionState.Starting:
                    _logger.LogDebug(
                        "Extension {ExtensionId} is already starting by another thread",
                        _definition.Id);
                    return false;
            }

            return await StartProcessAsync();
        }
        finally
        {
            _processLock.Release();
        }
    }

    /// <summary>
    ///     Performs the initialization handshake with the extension.
    ///     This method sends an "initialize" message and waits for "initialize_response"
    ///     containing the extension's metadata (name, version, etc.).
    /// </summary>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when the extension is not in a valid state for handshake,
    ///     or when the extension does not provide required metadata.
    /// </exception>
    /// <exception cref="OperationCanceledException">
    ///     Thrown when the handshake times out or is cancelled.
    /// </exception>
    public async Task PerformHandshakeAsync(CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (_state != ExtensionState.Starting &&
            _state != ExtensionState.Initializing &&
            _state != ExtensionState.Idle)
            throw new InvalidOperationException(
                $"Cannot perform handshake in state {_state}. " +
                "Extension must be Starting, Initializing, or Idle.");

        var process = _process;
        if (process == null || process.HasExited)
            throw new InvalidOperationException("Extension process is not running.");

        var initMessage = new
        {
            type = ExtensionMessageType.Initialize,
            protocolVersion = _definition.ProtocolVersion
        };
        var initJson = JsonSerializer.Serialize(initMessage);

        _logger.LogDebug("Sending initialize to extension {ExtensionId}", _definition.Id);

        if (!await WriteToStdinWithTimeoutAsync(process, initJson, cancellationToken))
            throw new InvalidOperationException("Failed to send initialize message to extension.");

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(TimeSpan.FromSeconds(_config.HandshakeTimeoutSeconds));

        var response = await WaitForMessageAsync<ExtensionInitializeResponse>(
            ExtensionMessageType.InitializeResponse,
            timeoutCts.Token);

        if (string.IsNullOrWhiteSpace(response.Name))
            throw new InvalidOperationException("Extension did not provide a name in handshake response.");

        if (string.IsNullOrWhiteSpace(response.Version))
            throw new InvalidOperationException("Extension did not provide a version in handshake response.");

        _definition.RuntimeMetadata = response;

        var initializedMessage = new { type = ExtensionMessageType.Initialized };
        var initializedJson = JsonSerializer.Serialize(initializedMessage);

        // ReSharper disable once PossiblyMistakenUseOfCancellationToken - Intentional: confirmation message uses external token, not timeout token
        if (!await WriteToStdinWithTimeoutAsync(process, initializedJson, cancellationToken))
            _logger.LogWarning(
                "Failed to send initialized confirmation to extension {ExtensionId}",
                _definition.Id);

        SetState(ExtensionState.Idle);

        _logger.LogDebug(
            "Handshake completed with extension {ExtensionId}: {Name} v{Version}",
            _definition.Id, response.Name, response.Version);
    }

    /// <summary>
    ///     Waits for the extension to complete initialization.
    ///     Returns when the extension transitions to Idle/Busy state or fails.
    /// </summary>
    /// <param name="timeout">
    ///     Optional timeout. Defaults to HandshakeTimeoutSeconds + 5 seconds.
    /// </param>
    /// <returns>True if initialization completed successfully, false otherwise.</returns>
    public async Task<bool> WaitForInitializationAsync(TimeSpan? timeout = null)
    {
        timeout ??= TimeSpan.FromSeconds(_config.HandshakeTimeoutSeconds + 5);
        var startTime = DateTime.UtcNow;

        while (DateTime.UtcNow - startTime < timeout)
        {
            var currentState = _state;

            if (currentState == ExtensionState.Idle || currentState == ExtensionState.Busy)
                return true;

            if (currentState == ExtensionState.Error ||
                currentState == ExtensionState.Crashed ||
                currentState == ExtensionState.Unloaded)
                return false;

            await Task.Delay(100);
        }

        _logger.LogWarning(
            "Timeout waiting for extension {ExtensionId} to complete initialization",
            _definition.Id);
        return false;
    }

    /// <summary>
    ///     Waits for a specific message type from the extension.
    /// </summary>
    /// <typeparam name="T">The type to deserialize the message to.</typeparam>
    /// <param name="expectedType">The expected message type.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>The deserialized message.</returns>
    private async Task<T> WaitForMessageAsync<T>(string expectedType, CancellationToken cancellationToken)
        where T : class
    {
        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);

        void OnMessage(object? sender, ExtensionMessageEventArgs e)
        {
            if (e.MessageType == expectedType)
                try
                {
                    var result = JsonSerializer.Deserialize<T>(e.RawJson);
                    if (result != null)
                        tcs.TrySetResult(result);
                    else
                        tcs.TrySetException(new InvalidOperationException(
                            $"Failed to deserialize {expectedType} message."));
                }
                catch (JsonException ex)
                {
                    tcs.TrySetException(new InvalidOperationException(
                        $"Failed to parse {expectedType} message: {ex.Message}", ex));
                }
        }

        MessageReceived += OnMessage;
        try
        {
            // ReSharper disable once UseAwaitUsing - CancellationTokenRegistration does not implement IAsyncDisposable
            using (cancellationToken.Register(() =>
                       tcs.TrySetCanceled(cancellationToken)))
            {
                return await tcs.Task;
            }
        }
        finally
        {
            MessageReceived -= OnMessage;
        }
    }

    /// <summary>
    ///     Sends a snapshot to the extension.
    /// </summary>
    /// <param name="data">Binary data to send.</param>
    /// <param name="metadata">Metadata describing the data.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>True if send was successful.</returns>
    public async Task<bool> SendSnapshotAsync(
        byte[] data,
        ExtensionMetadata metadata,
        CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return false;

        if (!await EnsureStartedAsync())
            return false;

        var generation = Interlocked.Read(ref _processGeneration);
        var process = _process;
        if (process == null || process.HasExited)
        {
            SetState(ExtensionState.Crashed);
            return false;
        }

        metadata.SequenceNumber = Interlocked.Increment(ref _sequenceNumber);
        metadata.Type = ExtensionMessageType.Snapshot;

        var activeCount = Interlocked.Increment(ref _activeOperations);
        if (activeCount == 1)
            TrySetBusyState();

        try
        {
            try
            {
                await _stdinLock.WaitAsync(cancellationToken);
            }
            catch (ObjectDisposedException)
            {
                return false;
            }

            bool success;
            try
            {
                var currentGeneration = Interlocked.Read(ref _processGeneration);
                if (currentGeneration != generation)
                {
                    _logger.LogDebug(
                        "Process was replaced during snapshot send for extension {ExtensionId}",
                        _definition.Id);
                    return false;
                }

                process = _process;

                if (Interlocked.Read(ref _processGeneration) != currentGeneration)
                {
                    _logger.LogDebug(
                        "Process was replaced while reading reference for extension {ExtensionId}",
                        _definition.Id);
                    return false;
                }

                if (process == null || process.HasExited)
                {
                    SetState(ExtensionState.Crashed);
                    return false;
                }

                try
                {
                    success = await _transport.SendAsync(process, data, metadata, cancellationToken);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex,
                        "Transport threw exception while sending snapshot to extension {ExtensionId}",
                        _definition.Id);
                    success = false;
                }
            }
            finally
            {
                _stdinLock.Release();
            }

            if (success)
            {
                _snapshotManager.RecordSnapshot(_definition.Id, metadata, EffectiveSnapshotTtlSeconds);
                Interlocked.Exchange(ref _lastActivityTicks, DateTime.UtcNow.Ticks);
                Interlocked.Exchange(ref _consecutiveFailures, 0);

                _logger.LogDebug(
                    "Sent snapshot to extension {ExtensionId}, sequence {SequenceNumber}, size {Size}",
                    _definition.Id, metadata.SequenceNumber, data.Length);
            }
            else
            {
                var failures = Interlocked.Increment(ref _consecutiveFailures);
                _logger.LogWarning(
                    "Failed to send snapshot to extension {ExtensionId}, consecutive failures: {Failures}",
                    _definition.Id, failures);

                if (_config.MaxConsecutiveSendFailures > 0 && failures >= _config.MaxConsecutiveSendFailures)
                {
                    _logger.LogError(
                        "Extension {ExtensionId} reached maximum consecutive send failures ({Max}), marking as crashed",
                        _definition.Id, _config.MaxConsecutiveSendFailures);
                    SetState(ExtensionState.Crashed);
                }
            }

            return success;
        }
        finally
        {
            var remaining = Interlocked.Decrement(ref _activeOperations);
            if (remaining == 0)
                TrySetIdleState();
        }
    }

    /// <summary>
    ///     Sends a heartbeat to the extension and checks for heartbeat timeout.
    ///     Uses dedicated heartbeat response tracking to avoid false positives from other activity.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>True if heartbeat was sent successfully and no timeout detected.</returns>
    public async Task<bool> SendHeartbeatAsync(CancellationToken cancellationToken = default)
    {
        var process = _process;
        if (_disposed || process == null || process.HasExited || _state != ExtensionState.Idle)
            return false;

        if (_definition.Capabilities?.SupportsHeartbeat != true)
            return true;

        try
        {
            await _stdinLock.WaitAsync(cancellationToken);
        }
        catch (ObjectDisposedException)
        {
            return false;
        }

        try
        {
            process = _process;
            if (process == null || process.HasExited)
                return false;

            var lastSent = Interlocked.Read(ref _lastHeartbeatSentTicks);
            var lastResponse = Interlocked.Read(ref _lastHeartbeatResponseTicks);

            if (lastSent > 0 && lastResponse < lastSent)
            {
                var missed = Interlocked.Increment(ref _missedHeartbeats);
                var currentResponse = Interlocked.Read(ref _lastHeartbeatResponseTicks);
                if (currentResponse >= lastSent)
                {
                    Interlocked.Decrement(ref _missedHeartbeats);
                }
                else if (missed >= EffectiveMaxMissedHeartbeats)
                {
                    _logger.LogWarning(
                        "Extension {ExtensionId} missed {Missed} heartbeats, marking as unresponsive",
                        _definition.Id, missed);
                    SetState(ExtensionState.Crashed);
                    return false;
                }
                else
                {
                    _logger.LogDebug(
                        "Extension {ExtensionId} missed heartbeat response ({Missed}/{Max})",
                        _definition.Id, missed, EffectiveMaxMissedHeartbeats);
                }
            }

            var heartbeat = new { type = ExtensionMessageType.Heartbeat };
            var json = JsonSerializer.Serialize(heartbeat);

            if (!await WriteToStdinWithTimeoutAsync(process, json, cancellationToken))
                return false;

            Interlocked.Exchange(ref _lastHeartbeatSentTicks, DateTime.UtcNow.Ticks);

            _logger.LogDebug("Sent heartbeat to extension {ExtensionId}", _definition.Id);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to send heartbeat to extension {ExtensionId}", _definition.Id);
            return false;
        }
        finally
        {
            _stdinLock.Release();
        }
    }

    /// <summary>
    ///     Sends a command to the extension and optionally waits for a response.
    /// </summary>
    /// <param name="sessionId">Session identifier for context.</param>
    /// <param name="commandType">Type of command to send.</param>
    /// <param name="payload">Command payload parameters.</param>
    /// <param name="waitForResponse">Whether to wait for a response from the extension.</param>
    /// <param name="timeoutMs">Timeout in milliseconds when waiting for response. Default is 30000.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>Command response if waitForResponse is true; otherwise a success response.</returns>
    public async Task<CommandResponse> SendCommandAsync(
        string sessionId,
        string commandType,
        Dictionary<string, object>? payload = null,
        bool waitForResponse = true,
        int timeoutMs = 30000,
        CancellationToken cancellationToken = default)
    {
        var process = _process;
        if (_disposed || process == null || process.HasExited)
            return CommandResponse.Failure("Extension is not running");

        if (_state is ExtensionState.Error or ExtensionState.Stopping or ExtensionState.Unloaded)
            return CommandResponse.Failure($"Extension is in {_state} state");

        var metadata = ExtensionMetadata.CreateCommand(sessionId, commandType, payload);
        var commandId = metadata.CommandId!;

        TaskCompletionSource<CommandResponse>? tcs = null;
        if (waitForResponse)
        {
            tcs = new TaskCompletionSource<CommandResponse>(TaskCreationOptions.RunContinuationsAsynchronously);
            _pendingCommands[commandId] = tcs;
        }

        try
        {
            await _stdinLock.WaitAsync(cancellationToken);
        }
        catch (ObjectDisposedException)
        {
            _pendingCommands.TryRemove(commandId, out _);
            return CommandResponse.Failure("Extension was disposed");
        }

        try
        {
            process = _process;
            if (process == null || process.HasExited)
            {
                _pendingCommands.TryRemove(commandId, out _);
                return CommandResponse.Failure("Extension process exited");
            }

            var json = JsonSerializer.Serialize(metadata);
            if (!await WriteToStdinWithTimeoutAsync(process, json, cancellationToken))
            {
                _pendingCommands.TryRemove(commandId, out _);
                return CommandResponse.Failure("Failed to write command to extension");
            }

            _logger.LogDebug(
                "Sent command {CommandType} (id: {CommandId}) to extension {ExtensionId}",
                commandType, commandId, _definition.Id);
        }
        catch (Exception ex)
        {
            _pendingCommands.TryRemove(commandId, out _);
            _logger.LogWarning(ex, "Failed to send command to extension {ExtensionId}", _definition.Id);
            return CommandResponse.Failure($"Failed to send command: {ex.Message}");
        }
        finally
        {
            _stdinLock.Release();
        }

        if (!waitForResponse)
            return CommandResponse.Success(commandId);

        try
        {
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(timeoutMs);

            var completedTask = await Task.WhenAny(tcs!.Task, Task.Delay(Timeout.Infinite, timeoutCts.Token));
            if (completedTask == tcs.Task)
                return await tcs.Task;

            _pendingCommands.TryRemove(commandId, out _);
            return CommandResponse.Failure("Command timed out waiting for response");
        }
        catch (OperationCanceledException)
        {
            _pendingCommands.TryRemove(commandId, out _);
            return CommandResponse.Failure("Command was cancelled");
        }
    }

    /// <summary>
    ///     Handles an acknowledgment message from the extension.
    /// </summary>
    /// <param name="sequenceNumber">Sequence number being acknowledged.</param>
    /// <param name="status">The processing status from the extension.</param>
    /// <param name="error">Optional error message if status indicates failure.</param>
    public void HandleAck(long sequenceNumber, string? status = null, string? error = null)
    {
        _snapshotManager.HandleAck(_definition.Id, sequenceNumber);
        Interlocked.Exchange(ref _lastActivityTicks, DateTime.UtcNow.Ticks);

        if (!string.IsNullOrEmpty(status) && !status.Equals("processed", StringComparison.OrdinalIgnoreCase))
            _logger.LogWarning(
                "Extension {ExtensionId} reported non-success status for sequence {SequenceNumber}: {Status}, Error: {Error}",
                _definition.Id, sequenceNumber, status, error ?? "(none)");
    }

    /// <summary>
    ///     Attempts to restart the extension after a crash.
    ///     Uses a two-phase approach: cleanup under lock, delay without lock, then start under lock.
    /// </summary>
    /// <returns>True if restart was successful.</returns>
    public async Task<bool> TryRestartAsync()
    {
        if (_disposed)
            return false;

        if (Interlocked.CompareExchange(ref _isRestarting, 1, 0) != 0)
        {
            _logger.LogDebug(
                "Extension {ExtensionId} restart already in progress, skipping",
                _definition.Id);
            return false;
        }

        try
        {
            try
            {
                await _processLock.WaitAsync();
            }
            catch (ObjectDisposedException)
            {
                return false;
            }

            try
            {
                if (Interlocked.CompareExchange(ref _restartCount, 0, 0) >= _config.MaxRestartAttempts)
                {
                    _logger.LogError(
                        "Extension {ExtensionId} exceeded maximum restart attempts ({Max})",
                        _definition.Id, _config.MaxRestartAttempts);
                    SetState(ExtensionState.Error);
                    return false;
                }

                var processStartTime = new DateTime(Interlocked.Read(ref _processStartTimeTicks), DateTimeKind.Utc);
                var uptime = DateTime.UtcNow - processStartTime;
                if (uptime.TotalSeconds < RapidCrashThresholdSeconds && processStartTime > DateTime.MinValue)
                {
                    var rapidCrashes = Interlocked.Increment(ref _rapidCrashCount);
                    _logger.LogWarning(
                        "Extension {ExtensionId} crashed rapidly after {Seconds:F1}s (rapid crash {Count}/{Max})",
                        _definition.Id, uptime.TotalSeconds, rapidCrashes, MaxRapidCrashes);

                    if (rapidCrashes >= MaxRapidCrashes)
                    {
                        _logger.LogError(
                            "Extension {ExtensionId} has {Count} consecutive rapid crashes, entering Error state. " +
                            "Check extension logs and configuration.",
                            _definition.Id, rapidCrashes);
                        SetState(ExtensionState.Error);
                        return false;
                    }
                }
                else
                {
                    var previousRapidCrashCount = Interlocked.Exchange(ref _rapidCrashCount, 0);
                    if (previousRapidCrashCount > 0)
                        _logger.LogDebug(
                            "Extension {ExtensionId} uptime exceeded rapid crash threshold ({Threshold}s), " +
                            "resetting rapid crash count from {PreviousCount} to 0",
                            _definition.Id, RapidCrashThresholdSeconds, previousRapidCrashCount);
                }

                await CleanupProcessAsync();

                var attempt = Interlocked.Increment(ref _restartCount);
                _logger.LogInformation(
                    "Restarting extension {ExtensionId}, attempt {Attempt}/{Max}",
                    _definition.Id, attempt, _config.MaxRestartAttempts);
            }
            finally
            {
                _processLock.Release();
            }

            await Task.Delay(TimeSpan.FromSeconds(_config.RestartCooldownSeconds));

            if (_disposed)
            {
                _logger.LogDebug(
                    "Extension {ExtensionId} was disposed during restart cooldown",
                    _definition.Id);
                return false;
            }

            try
            {
                await _processLock.WaitAsync();
            }
            catch (ObjectDisposedException)
            {
                return false;
            }

            try
            {
                if (_state == ExtensionState.Error || _state == ExtensionState.Stopping || _disposed)
                {
                    _logger.LogDebug(
                        "Extension {ExtensionId} state changed during restart cooldown, aborting restart",
                        _definition.Id);
                    return false;
                }

                return await StartProcessAsync();
            }
            finally
            {
                _processLock.Release();
            }
        }
        finally
        {
            Interlocked.Exchange(ref _isRestarting, 0);
        }
    }

    /// <summary>
    ///     Attempts to recover an extension from Error state.
    ///     Resets the restart counter and attempts to start the process.
    /// </summary>
    /// <returns>True if recovery was successful and extension is now running.</returns>
    public async Task<bool> TryRecoverFromErrorAsync()
    {
        if (_disposed)
            return false;

        try
        {
            await _processLock.WaitAsync();
        }
        catch (ObjectDisposedException)
        {
            return false;
        }

        try
        {
            if (_state != ExtensionState.Error)
            {
                _logger.LogDebug(
                    "Extension {ExtensionId} is not in Error state (current: {State}), cannot recover",
                    _definition.Id, _state);
                return false;
            }

            _logger.LogInformation(
                "Attempting to recover extension {ExtensionId} from Error state",
                _definition.Id);

            Interlocked.Exchange(ref _restartCount, 0);
            await CleanupProcessAsync();
            return await StartProcessAsync();
        }
        finally
        {
            _processLock.Release();
        }
    }

    /// <summary>
    ///     Stops the extension process.
    /// </summary>
    /// <param name="resetRestartCount">Whether to reset the restart counter (for normal shutdowns).</param>
    public async Task StopAsync(bool resetRestartCount = false)
    {
        // Note: Don't check _disposed here - StopAsync is called from DisposeAsync
        // and needs to complete even after _disposed is set

        try
        {
            await _processLock.WaitAsync();
        }
        catch (ObjectDisposedException)
        {
            return;
        }

        try
        {
            if (_process == null)
                return;

            SetState(ExtensionState.Stopping);

            if (resetRestartCount)
                Interlocked.Exchange(ref _restartCount, 0);

            if (_process.HasExited)
            {
                _logger.LogDebug(
                    "Extension {ExtensionId} already exited with code {ExitCode}, skipping graceful shutdown",
                    _definition.Id, _process.ExitCode);
            }
            else
            {
                var stdinLockAcquired = false;
                try
                {
                    try
                    {
                        await _stdinLock.WaitAsync();
                        stdinLockAcquired = true;
                    }
                    catch (ObjectDisposedException)
                    {
                        _process.Kill(true);
                    }

                    if (stdinLockAcquired)
                        try
                        {
                            if (!_process.HasExited)
                            {
                                var shutdown = new { type = ExtensionMessageType.Shutdown };
                                var json = JsonSerializer.Serialize(shutdown);

                                using var writeCts = new CancellationTokenSource(_config.StdinWriteTimeoutMs);
                                await _process.StandardInput.WriteLineAsync(json.AsMemory(), writeCts.Token);
                                await _process.StandardInput.FlushAsync(writeCts.Token);
                            }

                            using var cts =
                                new CancellationTokenSource(
                                    TimeSpan.FromSeconds(_config.GracefulShutdownTimeoutSeconds));
                            await _process.WaitForExitAsync(cts.Token);
                        }
                        finally
                        {
                            _stdinLock.Release();
                        }
                }
                catch (OperationCanceledException ex)
                {
                    _logger.LogWarning(ex, "Extension {ExtensionId} did not exit gracefully, killing process",
                        _definition.Id);
                    if (stdinLockAcquired)
                        try
                        {
                            _stdinLock.Release();
                        }
                        catch
                        {
                            // Ignore semaphore release errors during shutdown
                        }

                    _process.Kill(true);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Error during graceful shutdown of extension {ExtensionId}", _definition.Id);
                    if (stdinLockAcquired)
                        try
                        {
                            _stdinLock.Release();
                        }
                        catch
                        {
                            // Ignore semaphore release errors during shutdown
                        }

                    try
                    {
                        _process.Kill(true);
                    }
                    catch (Exception killEx)
                    {
                        _logger.LogError(killEx,
                            "Failed to kill extension process {ExtensionId} after graceful shutdown failure",
                            _definition.Id);
                    }
                }
            }

            await CleanupProcessAsync();
            SetState(ExtensionState.Unloaded);

            _logger.LogInformation("Extension {ExtensionId} stopped", _definition.Id);
        }
        finally
        {
            _processLock.Release();
        }
    }

    /// <summary>
    ///     Notifies the extension that a session has been closed.
    /// </summary>
    /// <param name="sessionId">Session identifier.</param>
    /// <param name="owner">Session owner information.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public async Task NotifySessionClosedAsync(
        string sessionId,
        SessionOwner? owner,
        CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return;

        var process = _process;
        if (process == null || process.HasExited)
            return;

        try
        {
            await _stdinLock.WaitAsync(cancellationToken);
        }
        catch (ObjectDisposedException)
        {
            return;
        }

        try
        {
            process = _process;
            if (process == null || process.HasExited)
                return;

            var message = new
            {
                type = ExtensionMessageType.SessionClosed,
                sessionId,
                owner
            };
            var json = JsonSerializer.Serialize(message);
            await WriteToStdinWithTimeoutAsync(process, json, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "Failed to notify session closed to extension {ExtensionId}",
                _definition.Id);
        }
        finally
        {
            _stdinLock.Release();
        }
    }

    /// <summary>
    ///     Notifies the extension that a session has been unbound (but session still exists).
    /// </summary>
    /// <param name="sessionId">Session identifier that was unbound.</param>
    /// <param name="owner">Session owner information.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public async Task NotifySessionUnboundAsync(
        string sessionId,
        SessionOwner? owner,
        CancellationToken cancellationToken = default)
    {
        if (_disposed)
            return;

        var process = _process;
        if (process == null || process.HasExited)
            return;

        try
        {
            await _stdinLock.WaitAsync(cancellationToken);
        }
        catch (ObjectDisposedException)
        {
            return;
        }

        try
        {
            process = _process;
            if (process == null || process.HasExited)
                return;

            var message = new
            {
                type = ExtensionMessageType.SessionUnbound,
                sessionId,
                owner
            };
            var json = JsonSerializer.Serialize(message);
            await WriteToStdinWithTimeoutAsync(process, json, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "Failed to notify session unbound to extension {ExtensionId}",
                _definition.Id);
        }
        finally
        {
            _stdinLock.Release();
        }
    }

    /// <summary>
    ///     Starts the extension process.
    /// </summary>
    /// <returns>
    ///     A task that resolves to <c>true</c> if the process started successfully; otherwise, <c>false</c>.
    /// </returns>
    private Task<bool> StartProcessAsync()
    {
        SetState(ExtensionState.Starting);

        try
        {
            var startInfo = CreateProcessStartInfo();
            _process = new Process { StartInfo = startInfo };
            _process.EnableRaisingEvents = true;
            _process.Exited += OnProcessExited;

            if (!_process.Start())
            {
                _logger.LogError("Failed to start extension {ExtensionId}", _definition.Id);
                _process.Exited -= OnProcessExited;
                _process.Dispose();
                _process = null;
                SetState(ExtensionState.Error);
                return Task.FromResult(false);
            }

            if (_process.HasExited)
            {
                _logger.LogWarning(
                    "Extension {ExtensionId} process exited immediately after start with code {ExitCode}",
                    _definition.Id, _process.ExitCode);
                _process.Exited -= OnProcessExited;
                _process.Dispose();
                _process = null;
                SetState(ExtensionState.Crashed);
                return Task.FromResult(false);
            }

            _processCleanupManager?.RegisterProcess(_process);
            _readerCts = new CancellationTokenSource();
            _stdoutReaderTask = ReadStdoutAsync(_process, _readerCts.Token);
            _stderrReaderTask = ReadStderrAsync(_process, _readerCts.Token);
            _snapshotManager.RegisterTransport(_definition.Id, _transport);

            Interlocked.Increment(ref _processGeneration);

            var startTime = DateTime.UtcNow.Ticks;
            Interlocked.Exchange(ref _processStartTimeTicks, startTime);
            Interlocked.Exchange(ref _lastActivityTicks, startTime);
            Interlocked.Exchange(ref _lastHeartbeatSentTicks, 0);
            Interlocked.Exchange(ref _lastHeartbeatResponseTicks, startTime);
            Interlocked.Exchange(ref _consecutiveFailures, 0);
            Interlocked.Exchange(ref _missedHeartbeats, 0);

            SetState(ExtensionState.Initializing);

            _logger.LogInformation(
                "Extension {ExtensionId} process started, PID={Pid}, awaiting handshake",
                _definition.Id, _process.Id);

            return Task.FromResult(true);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to start extension {ExtensionId}", _definition.Id);
            if (_process != null)
            {
                _process.Exited -= OnProcessExited;
                _process.Dispose();
                _process = null;
            }

            SetState(ExtensionState.Error);
            return Task.FromResult(false);
        }
    }

    /// <summary>
    ///     Creates process start information from the extension command configuration.
    ///     Uses ArgumentList for safe argument handling to prevent command injection.
    /// </summary>
    /// <returns>Configured <see cref="ProcessStartInfo" /> instance.</returns>
    private ProcessStartInfo CreateProcessStartInfo()
    {
        var command = _definition.Command;
        var resolution = ResolveCommand(command);

        var startInfo = new ProcessStartInfo
        {
            FileName = resolution.FileName,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardInputEncoding = new UTF8Encoding(false),
            StandardOutputEncoding = new UTF8Encoding(false),
            StandardErrorEncoding = new UTF8Encoding(false)
        };

        foreach (var arg in resolution.ArgumentList)
            startInfo.ArgumentList.Add(arg);

        if (!string.IsNullOrEmpty(command.WorkingDirectory))
            startInfo.WorkingDirectory = command.WorkingDirectory;

        if (command.Environment != null)
            foreach (var kvp in command.Environment)
                startInfo.EnvironmentVariables[kvp.Key] = kvp.Value;

        return startInfo;
    }

    /// <summary>
    ///     Resolves the command type to actual executable and argument list.
    ///     Returns arguments as a list for safe handling via ProcessStartInfo.ArgumentList.
    /// </summary>
    /// <param name="command">The extension command configuration.</param>
    /// <returns>A <see cref="CommandResolution" /> containing the file name and argument list.</returns>
    private static CommandResolution ResolveCommand(ExtensionCommand command)
    {
        var isWindows = OperatingSystem.IsWindows();
        var executable = command.Executable;
        var argumentList = ParseArguments(command.Arguments);

        return command.Type.ToLowerInvariant() switch
        {
            "node" => new CommandResolution(isWindows ? "node.exe" : "node",
                PrependExecutable(executable, argumentList)),
            "python" => new CommandResolution(isWindows ? "python.exe" : "python3",
                PrependExecutable(executable, argumentList)),
            "dotnet" => new CommandResolution("dotnet", PrependExecutable(executable, argumentList)),
            "npx" => ResolveNpxCommand(executable, argumentList),
            "pipx" => new CommandResolution(isWindows ? "pipx.exe" : "pipx",
                PrependRun(executable, argumentList)),
            _ => new CommandResolution(executable, argumentList)
        };
    }

    /// <summary>
    ///     Resolves npx command by directly executing npx-cli.js with node to avoid
    ///     shell script path resolution issues when executed via Process.Start.
    ///     Adds --yes flag to auto-confirm package installation since extensions
    ///     run non-interactively and cannot respond to prompts.
    /// </summary>
    /// <param name="packageName">The npm package name to execute.</param>
    /// <param name="arguments">Additional arguments for the package.</param>
    /// <returns>A <see cref="CommandResolution" /> for npx execution.</returns>
    private static CommandResolution ResolveNpxCommand(string packageName, List<string> arguments)
    {
        var npxArgs = new List<string>(arguments.Count + 2) { "--yes", packageName };
        npxArgs.AddRange(arguments);

        var isWindows = OperatingSystem.IsWindows();
        var npxCliPath = FindNpxCliPath(isWindows);

        if (npxCliPath != null)
        {
            var args = new List<string>(npxArgs.Count + 1) { npxCliPath };
            args.AddRange(npxArgs);
            return new CommandResolution(isWindows ? "node.exe" : "node", args);
        }

        return new CommandResolution(isWindows ? "npx.cmd" : "npx", npxArgs);
    }

    /// <summary>
    ///     Finds the path to npx-cli.js by searching PATH for node installation.
    /// </summary>
    /// <param name="isWindows">Whether the current platform is Windows.</param>
    /// <returns>The full path to npx-cli.js, or null if not found.</returns>
    private static string? FindNpxCliPath(bool isWindows)
    {
        var pathEnv = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrEmpty(pathEnv))
            return null;

        var nodeExeName = isWindows ? "node.exe" : "node";
        var paths = pathEnv.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries);

        foreach (var path in paths)
            try
            {
                var nodePath = Path.Combine(path, nodeExeName);
                if (!File.Exists(nodePath))
                    continue;

                var npxCliPath = Path.Combine(path, "node_modules", "npm", "bin", "npx-cli.js");
                if (File.Exists(npxCliPath))
                    return npxCliPath;
            }
            catch
            {
            }

        return null;
    }

    /// <summary>
    ///     Prepends the executable path to the argument list.
    /// </summary>
    /// <param name="executable">The executable path to prepend.</param>
    /// <param name="arguments">The existing argument list.</param>
    /// <returns>A new list with the executable prepended.</returns>
    private static List<string> PrependExecutable(string executable, List<string> arguments)
    {
        var result = new List<string>(arguments.Count + 1) { executable };
        result.AddRange(arguments);
        return result;
    }

    /// <summary>
    ///     Prepends "run" and the package name to the argument list (for pipx).
    /// </summary>
    /// <param name="packageName">The package name to run.</param>
    /// <param name="arguments">The existing argument list.</param>
    /// <returns>A new list with "run" and the package name prepended.</returns>
    private static List<string> PrependRun(string packageName, List<string> arguments)
    {
        var result = new List<string>(arguments.Count + 2) { "run", packageName };
        result.AddRange(arguments);
        return result;
    }

    /// <summary>
    ///     Parses an argument string into a list of individual arguments.
    ///     Handles quoted strings and escapes properly.
    /// </summary>
    /// <param name="argumentString">The argument string to parse.</param>
    /// <returns>A list of parsed arguments.</returns>
    /// <remarks>
    ///     <para>
    ///         This parser supports both single and double quotes for grouping arguments
    ///         containing spaces. Quotes themselves are not included in the output.
    ///     </para>
    ///     <para>Edge cases handled:</para>
    ///     <list type="bullet">
    ///         <item>Empty or whitespace-only input: returns empty list</item>
    ///         <item>Multiple consecutive spaces: treated as single separator</item>
    ///         <item>Unclosed quotes: remaining content treated as single argument</item>
    ///         <item>Mixed quote types: each quote type only closes itself</item>
    ///     </list>
    ///     <para>
    ///         Note: This parser does not enforce a maximum argument length or count.
    ///         Input validation should be performed at the configuration loading layer
    ///         via <see cref="ExtensionDefinition" /> validation.
    ///     </para>
    /// </remarks>
    private static List<string> ParseArguments(string? argumentString)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(argumentString))
            return result;

        var current = new StringBuilder();
        var inQuotes = false;
        var quoteChar = '\0';

        foreach (var c in argumentString)
            if (!inQuotes && (c == '"' || c == '\''))
            {
                inQuotes = true;
                quoteChar = c;
            }
            else if (inQuotes && c == quoteChar)
            {
                inQuotes = false;
                quoteChar = '\0';
            }
            else if (!inQuotes && char.IsWhiteSpace(c))
            {
                if (current.Length > 0)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
            }
            else
            {
                current.Append(c);
            }

        if (current.Length > 0)
            result.Add(current.ToString());

        return result;
    }

    /// <summary>
    ///     Reads stdout from the extension process and processes messages.
    ///     Detects unexpected stream closure and triggers crash state if needed.
    ///     Includes timeout protection to detect hung extensions.
    /// </summary>
    /// <param name="process">The extension process to read from.</param>
    /// <param name="cancellationToken">Cancellation token to stop reading.</param>
    /// <returns>A task representing the reading operation.</returns>
    private async Task ReadStdoutAsync(Process process, CancellationToken cancellationToken)
    {
        var readTimeout = TimeSpan.FromSeconds(_config.HealthCheckIntervalSeconds * 3);

        try
        {
            while (!cancellationToken.IsCancellationRequested && !process.HasExited)
            {
                string? line;
                try
                {
                    using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                    timeoutCts.CancelAfter(readTimeout);
                    line = await process.StandardOutput.ReadLineAsync(timeoutCts.Token);
                }
                catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
                {
                    if (process.HasExited)
                        break;
                    continue;
                }

                if (line == null)
                    break;

                ProcessMessage(line);
            }
        }
        catch (OperationCanceledException)
        {
            // Ignore cancellation during shutdown
        }
        catch (Exception ex)
        {
            if (!process.HasExited && !cancellationToken.IsCancellationRequested)
                _logger.LogWarning(ex, "Error reading stdout from extension {ExtensionId}", _definition.Id);
        }

        DetectUnexpectedStreamClosure(process);
    }

    /// <summary>
    ///     Reads stderr from the extension process and logs the output.
    ///     Includes timeout protection to prevent indefinite blocking.
    /// </summary>
    /// <param name="process">The extension process to read from.</param>
    /// <param name="cancellationToken">Cancellation token to stop reading.</param>
    /// <returns>A task representing the reading operation.</returns>
    private async Task ReadStderrAsync(Process process, CancellationToken cancellationToken)
    {
        var readTimeout = TimeSpan.FromSeconds(_config.HealthCheckIntervalSeconds * 3);

        try
        {
            while (!cancellationToken.IsCancellationRequested && !process.HasExited)
            {
                string? line;
                try
                {
                    using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                    timeoutCts.CancelAfter(readTimeout);
                    line = await process.StandardError.ReadLineAsync(timeoutCts.Token);
                }
                catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
                {
                    if (process.HasExited)
                        break;
                    continue;
                }

                if (line == null)
                    break;

                _logger.LogWarning("[{ExtensionId}:stderr] {Line}", _definition.Id, line);
            }
        }
        catch (OperationCanceledException)
        {
            // Ignore cancellation during shutdown
        }
        catch (Exception ex)
        {
            if (!process.HasExited && !cancellationToken.IsCancellationRequested)
                _logger.LogDebug(ex, "Error reading stderr from extension {ExtensionId}", _definition.Id);
        }
    }

    /// <summary>
    ///     Detects unexpected stream closure and triggers crash state if the process
    ///     exited while we were not in a stopping state.
    /// </summary>
    /// <param name="process">The extension process.</param>
    private void DetectUnexpectedStreamClosure(Process process)
    {
        if (_state is ExtensionState.Stopping or ExtensionState.Unloaded)
            return;

        if (process.HasExited && _state is ExtensionState.Idle or ExtensionState.Busy or ExtensionState.Starting)
        {
            _logger.LogWarning(
                "Extension {ExtensionId} stdout stream closed unexpectedly, process exit code: {ExitCode}",
                _definition.Id, process.ExitCode);
            SetState(ExtensionState.Crashed);
        }
    }

    /// <summary>
    ///     Processes an ACK message from the extension, extracting status and error information.
    /// </summary>
    /// <param name="root">The parsed JSON root element.</param>
    private void ProcessAckMessage(JsonElement root)
    {
        if (!root.TryGetProperty("sequenceNumber", out var seqElement))
        {
            _logger.LogWarning("Received ACK without sequenceNumber from extension {ExtensionId}", _definition.Id);
            return;
        }

        var sequenceNumber = seqElement.GetInt64();
        string? status = null;
        string? error = null;

        if (root.TryGetProperty("status", out var statusElement))
            status = statusElement.GetString();

        if (root.TryGetProperty("error", out var errorElement))
            error = errorElement.GetString();

        HandleAck(sequenceNumber, status, error);
    }

    /// <summary>
    ///     Processes a command result message from the extension.
    /// </summary>
    /// <param name="root">The parsed JSON root element.</param>
    private void ProcessCommandResultMessage(JsonElement root)
    {
        if (!root.TryGetProperty("commandId", out var commandIdElement))
        {
            _logger.LogWarning("Received command_result without commandId from extension {ExtensionId}",
                _definition.Id);
            return;
        }

        var commandId = commandIdElement.GetString();
        if (string.IsNullOrEmpty(commandId))
        {
            _logger.LogWarning("Received command_result with empty commandId from extension {ExtensionId}",
                _definition.Id);
            return;
        }

        if (!_pendingCommands.TryRemove(commandId, out var tcs))
        {
            _logger.LogDebug(
                "Received command_result for unknown command {CommandId} from extension {ExtensionId}",
                commandId, _definition.Id);
            return;
        }

        var isSuccess = true;
        string? error = null;
        Dictionary<string, object>? result = null;

        if (root.TryGetProperty("success", out var successElement))
            isSuccess = successElement.GetBoolean();

        if (root.TryGetProperty("error", out var errorElement))
            error = errorElement.GetString();

        if (root.TryGetProperty("result", out var resultElement) && resultElement.ValueKind == JsonValueKind.Object)
        {
            result = new Dictionary<string, object>();
            foreach (var prop in resultElement.EnumerateObject()) result[prop.Name] = GetJsonValue(prop.Value);
        }

        var response = isSuccess
            ? CommandResponse.Success(commandId, result)
            : CommandResponse.Failure(error ?? "Unknown error", commandId);

        tcs.TrySetResult(response);

        Interlocked.Exchange(ref _lastActivityTicks, DateTime.UtcNow.Ticks);
        _logger.LogDebug(
            "Received command_result for {CommandId} from extension {ExtensionId}: success={Success}",
            commandId, _definition.Id, isSuccess);
    }

    /// <summary>
    ///     Converts a JsonElement to an appropriate .NET object.
    /// </summary>
    /// <param name="element">The JSON element to convert.</param>
    /// <returns>The converted value.</returns>
    private static object GetJsonValue(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString() ?? string.Empty,
            JsonValueKind.Number => element.TryGetInt64(out var l) ? l : element.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null!,
            JsonValueKind.Array => element.EnumerateArray().Select(GetJsonValue).ToList(),
            JsonValueKind.Object => element.EnumerateObject().ToDictionary(p => p.Name, p => GetJsonValue(p.Value)),
            _ => element.GetRawText()
        };
    }

    /// <summary>
    ///     Processes a JSON message received from the extension.
    /// </summary>
    /// <param name="line">The raw JSON line to process.</param>
    private void ProcessMessage(string line)
    {
        if (string.IsNullOrWhiteSpace(line) || _disposed)
            return;

        try
        {
            using var doc = JsonDocument.Parse(line);
            var root = doc.RootElement;

            if (!root.TryGetProperty("type", out var typeElement))
                return;

            var messageType = typeElement.GetString();

            switch (messageType)
            {
                case ExtensionMessageType.Ack:
                    ProcessAckMessage(root);
                    break;

                case ExtensionMessageType.Pong:
                    var pongTime = DateTime.UtcNow.Ticks;
                    Interlocked.Exchange(ref _lastActivityTicks, pongTime);
                    Interlocked.Exchange(ref _lastHeartbeatResponseTicks, pongTime);
                    Interlocked.Exchange(ref _missedHeartbeats, 0);
                    _logger.LogDebug("Received pong from extension {ExtensionId}", _definition.Id);
                    break;

                case ExtensionMessageType.CommandResult:
                    ProcessCommandResultMessage(root);
                    break;

                default:
                    try
                    {
                        MessageReceived?.Invoke(this, new ExtensionMessageEventArgs(messageType ?? "unknown", line));
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex,
                            "Event subscriber threw exception for message from extension {ExtensionId}",
                            _definition.Id);
                    }

                    break;
            }
        }
        catch (JsonException ex)
        {
            _logger.LogWarning(ex, "Failed to parse message from extension {ExtensionId}: {Line}",
                _definition.Id, line);
        }
    }

    /// <summary>
    ///     Event handler called when the extension process exits unexpectedly.
    /// </summary>
    /// <param name="sender">The event sender.</param>
    /// <param name="e">Event arguments.</param>
    private void OnProcessExited(object? sender, EventArgs e)
    {
        try
        {
            if (_state is ExtensionState.Stopping or ExtensionState.Unloaded or ExtensionState.Starting)
                return;

            var exitCode = -1;
            try
            {
                exitCode = _process?.ExitCode ?? -1;
            }
            catch (InvalidOperationException)
            {
                // Ignore: process may have already been disposed
            }

            _logger.LogWarning(
                "Extension {ExtensionId} exited unexpectedly with code {ExitCode}",
                _definition.Id, exitCode);

            SetState(ExtensionState.Crashed);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex,
                "Unhandled exception in OnProcessExited for extension {ExtensionId}",
                _definition.Id);
        }
    }

    /// <summary>
    ///     Cleans up the process resources and unregisters from snapshot manager.
    ///     Ensures reader tasks complete before disposing process to prevent access violations.
    /// </summary>
    /// <returns>A task representing the cleanup operation.</returns>
    private async Task CleanupProcessAsync()
    {
        _snapshotManager.CleanupExtensionSnapshots(_definition.Id);
        _snapshotManager.UnregisterTransport(_definition.Id);

        if (_readerCts != null)
            try
            {
                await _readerCts.CancelAsync();
            }
            catch
            {
                // Ignore cancellation errors during cleanup
            }

        var readerTaskTimeout = TimeSpan.FromSeconds(ReaderTaskCleanupTimeoutSeconds);

        if (_stdoutReaderTask != null)
        {
            try
            {
                await _stdoutReaderTask.WaitAsync(readerTaskTimeout);
            }
            catch (TimeoutException ex)
            {
                _logger.LogWarning(ex,
                    "stdout reader task for extension {ExtensionId} did not complete within {Timeout}s. " +
                    "This may indicate the process is hung or streams are blocked.",
                    _definition.Id, ReaderTaskCleanupTimeoutSeconds);
            }
            catch
            {
                // Ignore reader task errors during cleanup
            }

            _stdoutReaderTask = null;
        }

        if (_stderrReaderTask != null)
        {
            try
            {
                await _stderrReaderTask.WaitAsync(readerTaskTimeout);
            }
            catch (TimeoutException ex)
            {
                _logger.LogWarning(ex,
                    "stderr reader task for extension {ExtensionId} did not complete within {Timeout}s",
                    _definition.Id, ReaderTaskCleanupTimeoutSeconds);
            }
            catch
            {
                // Ignore reader task errors during cleanup
            }

            _stderrReaderTask = null;
        }

        if (_readerCts != null)
        {
            _readerCts.Dispose();
            _readerCts = null;
        }

        if (_process != null)
        {
            _processCleanupManager?.UnregisterProcess(_process);
            _process.Exited -= OnProcessExited;
            _process.Dispose();
            _process = null;
        }
    }

    /// <summary>
    ///     Attempts to transition from Idle to Busy state.
    ///     Only succeeds if current state is Idle.
    /// </summary>
    /// <remarks>
    ///     This method is called when the first active operation starts.
    ///     State transition may fail if the extension is not in Idle state
    ///     (e.g., during shutdown or after crash). This is expected and logged at Trace level.
    /// </remarks>
    private void TrySetBusyState()
    {
        lock (_stateLock)
        {
            var currentState = _state;
            if (currentState == ExtensionState.Idle)
            {
                _state = ExtensionState.Busy;
            }
            else
            {
                _logger.LogTrace(
                    "Extension {ExtensionId} cannot transition to Busy: current state is {CurrentState}",
                    _definition.Id, currentState);
                return;
            }
        }

        _logger.LogDebug(
            "Extension {ExtensionId} state changed: Idle -> Busy",
            _definition.Id);

        try
        {
            StateChanged?.Invoke(this, ExtensionState.Busy);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "StateChanged event subscriber threw exception for extension {ExtensionId}",
                _definition.Id);
        }
    }

    /// <summary>
    ///     Attempts to transition from Busy to Idle state.
    ///     Only succeeds if current state is Busy and no active operations.
    /// </summary>
    /// <remarks>
    ///     This method is called when the last active operation completes.
    ///     State transition may fail if:
    ///     - The extension is not in Busy state (e.g., crashed during operation)
    ///     - There are still active operations (concurrent sends started during this call)
    ///     These cases are expected and logged at Trace level.
    /// </remarks>
    private void TrySetIdleState()
    {
        lock (_stateLock)
        {
            var currentState = _state;
            var activeOps = Interlocked.CompareExchange(ref _activeOperations, 0, 0);
            if (currentState == ExtensionState.Busy && activeOps == 0)
            {
                _state = ExtensionState.Idle;
            }
            else
            {
                _logger.LogTrace(
                    "Extension {ExtensionId} cannot transition to Idle: state={CurrentState}, activeOperations={ActiveOps}",
                    _definition.Id, currentState, activeOps);
                return;
            }
        }

        _logger.LogDebug(
            "Extension {ExtensionId} state changed: Busy -> Idle",
            _definition.Id);

        try
        {
            StateChanged?.Invoke(this, ExtensionState.Idle);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "StateChanged event subscriber threw exception for extension {ExtensionId}",
                _definition.Id);
        }
    }

    /// <summary>
    ///     Sets the extension state and raises the <see cref="StateChanged" /> event.
    ///     Thread-safe via lock.
    /// </summary>
    /// <param name="newState">The new state to set.</param>
    private void SetState(ExtensionState newState)
    {
        ExtensionState oldState;

        lock (_stateLock)
        {
            if (_state == newState)
                return;

            oldState = _state;
            _state = newState;
        }

        _logger.LogDebug(
            "Extension {ExtensionId} state changed: {OldState} -> {NewState}",
            _definition.Id, oldState, newState);

        try
        {
            StateChanged?.Invoke(this, newState);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "StateChanged event subscriber threw exception for extension {ExtensionId}",
                _definition.Id);
        }
    }

    /// <summary>
    ///     Writes a JSON message to the extension's stdin with timeout protection.
    ///     Prevents indefinite blocking if the extension's stdin buffer is full.
    /// </summary>
    /// <param name="process">The extension process to write to.</param>
    /// <param name="json">The JSON message to write.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns><c>true</c> if the write completed successfully; otherwise, <c>false</c>.</returns>
    private async Task<bool> WriteToStdinWithTimeoutAsync(
        Process process,
        string json,
        CancellationToken cancellationToken)
    {
        try
        {
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(_config.StdinWriteTimeoutMs);

            await process.StandardInput.WriteLineAsync(json.AsMemory(), timeoutCts.Token);
            await process.StandardInput.FlushAsync(timeoutCts.Token);
            return true;
        }
        catch (OperationCanceledException ex) when (!cancellationToken.IsCancellationRequested)
        {
            _logger.LogWarning(ex,
                "Stdin write to extension {ExtensionId} timed out after {Timeout}ms",
                _definition.Id, _config.StdinWriteTimeoutMs);
            return false;
        }
    }
}
