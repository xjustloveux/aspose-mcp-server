using System.Collections.Concurrent;
using AsposeMcpServer.Core.Extension.Transport;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Manages snapshot lifecycle, ensuring proper resource cleanup.
///     Handles ACK tracking and TTL-based cleanup for pending snapshots.
/// </summary>
// ReSharper disable once ClassWithVirtualMembersNeverInherited.Global - Virtual Dispose(bool) for proper IDisposable pattern
public class SnapshotManager : IHostedService, IDisposable
{
    /// <summary>
    ///     Maximum number of pending snapshots to prevent unbounded memory growth.
    ///     When exceeded, oldest snapshots are evicted.
    /// </summary>
    /// <remarks>
    ///     10,000 snapshots at ~1KB metadata each = ~10MB overhead.
    ///     Actual memory usage is higher if file transport stores data bytes.
    ///     This limit provides reasonable capacity for high-throughput scenarios
    ///     while preventing runaway memory consumption from unresponsive extensions.
    /// </remarks>
    private const int MaxPendingSnapshots = 10000;

    /// <summary>
    ///     Percentage of snapshots to evict when limit is reached.
    /// </summary>
    /// <remarks>
    ///     10% eviction (1,000 snapshots) provides a buffer so we don't
    ///     immediately hit the limit again after eviction.
    ///     Larger values reduce eviction frequency but may discard more data.
    /// </remarks>
    private const double EvictionPercentage = 0.1;

    /// <summary>
    ///     The extension configuration containing TTL and other settings.
    /// </summary>
    private readonly ExtensionConfig _config;

    /// <summary>
    ///     Lock object for eviction operations.
    /// </summary>
    private readonly object _evictionLock = new();

    /// <summary>
    ///     Logger instance for diagnostic output.
    /// </summary>
    private readonly ILogger<SnapshotManager> _logger;

    /// <summary>
    ///     Dictionary of pending snapshots keyed by "extensionId:sequenceNumber".
    /// </summary>
    private readonly ConcurrentDictionary<string, SnapshotRecord> _pendingSnapshots = new();

    /// <summary>
    ///     Dictionary of registered transports keyed by extension ID.
    /// </summary>
    private readonly ConcurrentDictionary<string, IExtensionTransport> _transports = new();

    /// <summary>
    ///     Cancellation token source for the cleanup loop.
    /// </summary>
    private CancellationTokenSource? _cleanupCts;

    /// <summary>
    ///     The background cleanup task.
    /// </summary>
    private Task? _cleanupTask;

    /// <summary>
    ///     Whether this instance has been disposed.
    /// </summary>
    private bool _disposed;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SnapshotManager" /> class.
    /// </summary>
    /// <param name="config">Extension configuration.</param>
    /// <param name="logger">Logger instance.</param>
    public SnapshotManager(ExtensionConfig config, ILogger<SnapshotManager> logger)
    {
        _config = config;
        _logger = logger;
    }

    /// <summary>
    ///     Gets the count of pending (unacknowledged) snapshots.
    /// </summary>
    public int PendingSnapshotCount => _pendingSnapshots.Count;

    /// <inheritdoc />
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <inheritdoc />
    public Task StartAsync(CancellationToken cancellationToken)
    {
        if (!_config.Enabled)
            return Task.CompletedTask;

        _cleanupCts = new CancellationTokenSource();
        _cleanupTask = RunCleanupLoopAsync(_cleanupCts.Token);

        _logger.LogInformation("SnapshotManager started with TTL={TtlSeconds}s", _config.SnapshotTtlSeconds.Default);
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    public async Task StopAsync(CancellationToken cancellationToken)
    {
        if (_cleanupCts == null)
            return;

        await _cleanupCts.CancelAsync();

        if (_cleanupTask != null)
            try
            {
                await _cleanupTask.WaitAsync(cancellationToken);
            }
            catch (OperationCanceledException)
            {
                // Ignore cancellation during shutdown
            }

        CleanupAllPendingSnapshots();
        _logger.LogInformation("SnapshotManager stopped");
    }

    /// <summary>
    ///     Registers a transport for a specific extension.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    /// <param name="transport">Transport instance.</param>
    public void RegisterTransport(string extensionId, IExtensionTransport transport)
    {
        _transports[extensionId] = transport;
    }

    /// <summary>
    ///     Unregisters a transport for a specific extension.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    public void UnregisterTransport(string extensionId)
    {
        _transports.TryRemove(extensionId, out _);
    }

    /// <summary>
    ///     Records a snapshot that was sent to an extension.
    ///     Enforces a limit on pending snapshots to prevent unbounded memory growth.
    /// </summary>
    /// <remarks>
    ///     Note: There is a potential race condition between the count check and eviction.
    ///     Multiple threads could pass the limit check simultaneously before any eviction occurs.
    ///     This is intentionally allowed as:
    ///     1. The lock in EvictOldestSnapshots re-checks the count and handles concurrent entries.
    ///     2. Slightly exceeding the limit temporarily is acceptable for performance.
    ///     3. The limit serves as a soft cap to prevent unbounded growth, not a hard boundary.
    /// </remarks>
    /// <param name="extensionId">Extension identifier.</param>
    /// <param name="metadata">Snapshot metadata.</param>
    /// <param name="ttlSeconds">
    ///     Optional TTL in seconds for this snapshot.
    ///     If not specified, uses global config value.
    /// </param>
    public void RecordSnapshot(string extensionId, ExtensionMetadata metadata, int? ttlSeconds = null)
    {
        if (_pendingSnapshots.Count >= MaxPendingSnapshots)
            EvictOldestSnapshots();

        var key = GetSnapshotKey(extensionId, metadata.SequenceNumber);
        var record = new SnapshotRecord
        {
            ExtensionId = extensionId,
            Metadata = metadata,
            SentAt = DateTime.UtcNow,
            TtlSeconds = ttlSeconds ?? _config.SnapshotTtlSeconds.Default
        };
        _pendingSnapshots[key] = record;

        _logger.LogDebug(
            "Recorded snapshot for extension {ExtensionId}, sequence {SequenceNumber}",
            extensionId, metadata.SequenceNumber);
    }

    /// <summary>
    ///     Evicts oldest pending snapshots when the limit is reached.
    ///     Cleans up associated resources for evicted snapshots.
    /// </summary>
    /// <remarks>
    ///     <para>Eviction strategy:</para>
    ///     <list type="bullet">
    ///         <item>Uses lock to prevent concurrent eviction attempts</item>
    ///         <item>Re-checks count after acquiring lock (another thread may have evicted)</item>
    ///         <item>Evicts oldest snapshots first (FIFO-like based on SentAt)</item>
    ///         <item>Cleans up transport resources for each evicted snapshot</item>
    ///     </list>
    ///     <para>
    ///         Evicted snapshots are logged at Warning level since this indicates
    ///         extensions may not be acknowledging snapshots properly.
    ///     </para>
    /// </remarks>
    private void EvictOldestSnapshots()
    {
        lock (_evictionLock)
        {
            if (_pendingSnapshots.Count < MaxPendingSnapshots)
                return;

            var countToEvict = (int)(MaxPendingSnapshots * EvictionPercentage);
            countToEvict = Math.Max(countToEvict, 1);

            var toEvict = _pendingSnapshots
                .OrderBy(kvp => kvp.Value.SentAt)
                .Take(countToEvict)
                .Select(kvp => kvp.Key)
                .ToList();

            foreach (var key in toEvict)
                if (_pendingSnapshots.TryRemove(key, out var record))
                {
                    CleanupSnapshotResources(record);
                    _logger.LogWarning(
                        "Evicted pending snapshot for extension {ExtensionId}, sequence {SequenceNumber} " +
                        "due to limit ({Limit}). Extension may not be acknowledging snapshots.",
                        record.ExtensionId, record.Metadata.SequenceNumber, MaxPendingSnapshots);
                }
        }
    }

    /// <summary>
    ///     Handles acknowledgment from an extension, cleaning up the associated snapshot.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    /// <param name="sequenceNumber">Sequence number being acknowledged.</param>
    /// <returns>True if the snapshot was found and cleaned up; otherwise, false.</returns>
    public bool HandleAck(string extensionId, long sequenceNumber)
    {
        var key = GetSnapshotKey(extensionId, sequenceNumber);

        if (!_pendingSnapshots.TryRemove(key, out var record))
        {
            _logger.LogWarning(
                "Received ack for unknown snapshot: extension {ExtensionId}, sequence {SequenceNumber}",
                extensionId, sequenceNumber);
            return false;
        }

        CleanupSnapshotResources(record);

        var duration = DateTime.UtcNow - record.SentAt;
        _logger.LogDebug(
            "Acknowledged snapshot for extension {ExtensionId}, sequence {SequenceNumber}, duration {Duration}ms",
            extensionId, sequenceNumber, duration.TotalMilliseconds);

        return true;
    }

    /// <summary>
    ///     Gets the count of pending snapshots for a specific extension.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    /// <returns>Count of pending snapshots.</returns>
    public int GetPendingSnapshotCount(string extensionId)
    {
        return _pendingSnapshots.Values.Count(r => r.ExtensionId == extensionId);
    }

    /// <summary>
    ///     Cleans up all pending snapshots for a specific extension.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    public void CleanupExtensionSnapshots(string extensionId)
    {
        var keysToRemove = _pendingSnapshots
            .Where(kvp => kvp.Value.ExtensionId == extensionId)
            .Select(kvp => kvp.Key)
            .ToList();

        foreach (var key in keysToRemove)
            if (_pendingSnapshots.TryRemove(key, out var record))
                CleanupSnapshotResources(record);

        if (keysToRemove.Count > 0)
            _logger.LogDebug(
                "Cleaned up {Count} pending snapshots for extension {ExtensionId}",
                keysToRemove.Count, extensionId);
    }

    /// <summary>
    ///     Disposes resources.
    /// </summary>
    /// <param name="disposing">
    ///     <c>true</c> to dispose managed resources; <c>false</c> to dispose only unmanaged resources.
    /// </param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing)
        {
            _cleanupCts?.Cancel();
            _cleanupCts?.Dispose();
            CleanupAllPendingSnapshots();
        }

        _disposed = true;
    }

    /// <summary>
    ///     Runs the background cleanup loop that removes expired snapshots.
    /// </summary>
    /// <param name="cancellationToken">Token to cancel the loop.</param>
    /// <returns>A task representing the cleanup loop.</returns>
    private async Task RunCleanupLoopAsync(CancellationToken cancellationToken)
    {
        var checkInterval = TimeSpan.FromSeconds(Math.Max(1, _config.SnapshotTtlSeconds.Default / 2));

        while (!cancellationToken.IsCancellationRequested)
            try
            {
                await Task.Delay(checkInterval, cancellationToken);
                CleanupExpiredSnapshots();
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in snapshot cleanup loop");
            }
    }

    /// <summary>
    ///     Cleans up all snapshots that have exceeded their TTL.
    /// </summary>
    /// <remarks>
    ///     Called periodically by the cleanup loop. Snapshots exceeding TTL
    ///     are assumed to be from unresponsive extensions and are cleaned up
    ///     to prevent resource leaks.
    /// </remarks>
    private void CleanupExpiredSnapshots()
    {
        var now = DateTime.UtcNow;
        var expiredKeys = new List<string>();

        foreach (var kvp in _pendingSnapshots)
        {
            var recordTtl = TimeSpan.FromSeconds(kvp.Value.TtlSeconds);
            if (now - kvp.Value.SentAt > recordTtl)
                expiredKeys.Add(kvp.Key);
        }

        var cleanedCount = 0;
        foreach (var key in expiredKeys)
            if (_pendingSnapshots.TryRemove(key, out var record))
            {
                _logger.LogWarning(
                    "Snapshot TTL expired for extension {ExtensionId}, sequence {SequenceNumber}. " +
                    "Extension may not be responding to acks.",
                    record.ExtensionId, record.Metadata.SequenceNumber);

                CleanupSnapshotResources(record);
                cleanedCount++;
            }

        if (cleanedCount > 0)
            _logger.LogDebug(
                "TTL cleanup completed: {CleanedCount} expired snapshot(s) removed, {RemainingCount} pending",
                cleanedCount, _pendingSnapshots.Count);
    }

    /// <summary>
    ///     Cleans up all pending snapshots regardless of TTL.
    ///     Called during shutdown.
    ///     Uses iterative approach to handle entries added during cleanup.
    /// </summary>
    private void CleanupAllPendingSnapshots()
    {
        while (!_pendingSnapshots.IsEmpty)
            foreach (var kvp in _pendingSnapshots)
                if (_pendingSnapshots.TryRemove(kvp.Key, out var record))
                    CleanupSnapshotResources(record);
    }

    /// <summary>
    ///     Cleans up the transport resources for a snapshot record.
    ///     Falls back to direct file cleanup if transport is not registered.
    /// </summary>
    /// <param name="record">The snapshot record to clean up.</param>
    private void CleanupSnapshotResources(SnapshotRecord record)
    {
        if (_transports.TryGetValue(record.ExtensionId, out var transport))
            try
            {
                transport.Cleanup(record.Metadata);
                return;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex,
                    "Failed to cleanup snapshot resources for extension {ExtensionId}, sequence {SequenceNumber}",
                    record.ExtensionId, record.Metadata.SequenceNumber);
            }

        CleanupSnapshotFileFallback(record);
    }

    /// <summary>
    ///     Directly cleans up the snapshot file when transport is not available.
    ///     This is a fallback for cases where the transport was unregistered before TTL expiration.
    ///     Note: mmap resources cannot be cleaned up without the transport instance and will be
    ///     cleaned when the MmapTransport is disposed.
    /// </summary>
    /// <param name="record">The snapshot record to clean up.</param>
    private void CleanupSnapshotFileFallback(SnapshotRecord record)
    {
        var filePath = record.Metadata.FilePath;
        var mmapName = record.Metadata.MmapName;

        if (!string.IsNullOrEmpty(filePath))
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    _logger.LogDebug(
                        "Cleaned up orphaned snapshot file for extension {ExtensionId}, sequence {SequenceNumber}: {FilePath}",
                        record.ExtensionId, record.Metadata.SequenceNumber, filePath);
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex,
                    "Failed to cleanup orphaned snapshot file: {FilePath}",
                    filePath);
            }
        else if (!string.IsNullOrEmpty(mmapName))
            _logger.LogDebug(
                "Orphaned mmap snapshot for extension {ExtensionId}, sequence {SequenceNumber} " +
                "will be cleaned when transport is disposed: {MmapName}",
                record.ExtensionId, record.Metadata.SequenceNumber, mmapName);
    }

    /// <summary>
    ///     Generates a unique key for a snapshot based on extension ID and sequence number.
    /// </summary>
    /// <param name="extensionId">The extension identifier.</param>
    /// <param name="sequenceNumber">The sequence number of the snapshot.</param>
    /// <returns>A composite key string in the format "extensionId:sequenceNumber".</returns>
    private static string GetSnapshotKey(string extensionId, long sequenceNumber)
    {
        return $"{extensionId}:{sequenceNumber}";
    }

    /// <summary>
    ///     Record of a pending snapshot awaiting acknowledgment.
    /// </summary>
    /// <remarks>
    ///     <para>Lifecycle:</para>
    ///     <list type="number">
    ///         <item>Created when <see cref="RecordSnapshot" /> is called after sending</item>
    ///         <item>Removed when <see cref="HandleAck" /> receives acknowledgment</item>
    ///         <item>
    ///             Removed by TTL cleanup if not acknowledged within <see cref="ExtensionConfig.SnapshotTtlSeconds" />
    ///         </item>
    ///         <item>May be evicted early if <see cref="MaxPendingSnapshots" /> is exceeded</item>
    ///     </list>
    /// </remarks>
    private sealed class SnapshotRecord
    {
        /// <summary>
        ///     Gets the extension identifier that received this snapshot.
        /// </summary>
        public required string ExtensionId { get; init; }

        /// <summary>
        ///     Gets the metadata associated with this snapshot.
        /// </summary>
        public required ExtensionMetadata Metadata { get; init; }

        /// <summary>
        ///     Gets the UTC timestamp when this snapshot was sent.
        /// </summary>
        public required DateTime SentAt { get; init; }

        /// <summary>
        ///     Gets the TTL in seconds for this snapshot.
        ///     Extension-specific TTL takes precedence over global config.
        /// </summary>
        public required int TtlSeconds { get; init; }
    }
}
