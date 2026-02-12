using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Hashing;
using System.IO.MemoryMappedFiles;
using System.Text.Json;

namespace AsposeMcpServer.Core.Extension.Transport;

/// <summary>
///     Memory-mapped file transport for extensions.
///     Uses shared memory for high-performance data transfer.
///     Supports cross-platform operation:
///     - Windows: Named shared memory via kernel objects
///     - Linux: POSIX shm_open via /dev/shm/
///     - macOS: File-backed memory mapping (automatic fallback)
/// </summary>
// ReSharper disable once ClassWithVirtualMembersNeverInherited.Global - Virtual Dispose(bool) for proper IDisposable pattern
public class MmapTransport : IExtensionTransport, IDisposable
{
    /// <summary>
    ///     Maximum number of active memory-mapped files to prevent file descriptor exhaustion.
    /// </summary>
    private const int MaxActiveMmaps = 500;

    /// <summary>
    ///     Percentage of mmaps to evict when limit is reached to prevent frequent eviction.
    /// </summary>
    /// <remarks>
    ///     10% eviction (50 mmaps when at 500 limit) provides a balance between:
    ///     - Not evicting too aggressively (extensions need time to read data)
    ///     - Not evicting too little (would cause frequent eviction cycles)
    ///     Evicted mmaps go to delayed cleanup queue for a final grace period.
    /// </remarks>
    private const double EvictionPercentage = 0.1;

    /// <summary>
    ///     Default maximum data size in bytes (100 MB).
    /// </summary>
    private const long DefaultMaxDataSize = 100 * 1024 * 1024;

    /// <summary>
    ///     Delay in milliseconds before disposing an mmap after cleanup request.
    ///     Gives child process time to finish reading the data.
    /// </summary>
    /// <remarks>
    ///     <para>
    ///         1 second delay provides reasonable time for extension to:
    ///         - Complete any in-progress read operation
    ///         - Process the acknowledgment
    ///         - Handle any I/O buffering
    ///     </para>
    ///     <para>
    ///         If extensions consistently fail to read data before cleanup,
    ///         this value can be increased, but doing so increases memory pressure.
    ///     </para>
    /// </remarks>
    private const int CleanupDelayMs = 1000;

    /// <summary>
    ///     Maximum number of mmaps in the pending cleanup queue.
    ///     When exceeded, oldest pending items are force-disposed immediately.
    /// </summary>
    private const int MaxPendingCleanup = 1000;

    /// <summary>
    ///     The mmap creation strategy for the current platform.
    /// </summary>
    private static readonly MmapStrategy CurrentStrategy = DetermineStrategy();

    /// <summary>
    ///     Counter for generating unique mmap names within the same process.
    /// </summary>
    private static long _mmapCounter;

    /// <summary>
    ///     Dictionary of active memory-mapped files keyed by their names, with creation timestamps.
    /// </summary>
    private readonly ConcurrentDictionary<string, MmapEntry> _activeMmaps = new();

    /// <summary>
    ///     Directory for file-backed mmap storage (macOS and other non-Windows/Linux platforms).
    /// </summary>
    private readonly string? _backingFileDirectory;

    /// <summary>
    ///     Timer for processing delayed cleanup queue.
    /// </summary>
    private readonly Timer _cleanupTimer;

    /// <summary>
    ///     Lock object for eviction operations.
    /// </summary>
    private readonly object _evictionLock = new();

    /// <summary>
    ///     Logger instance for diagnostic output.
    /// </summary>
    private readonly ILogger<MmapTransport>? _logger;

    /// <summary>
    ///     Maximum data size in bytes.
    /// </summary>
    private readonly long _maxDataSize;

    /// <summary>
    ///     Queue of mmaps pending delayed cleanup.
    /// </summary>
    private readonly ConcurrentQueue<(string Name, MmapEntry Entry, DateTime ScheduledAt)> _pendingCleanup = new();

    /// <summary>
    ///     Whether this instance has been disposed.
    /// </summary>
    private volatile bool _disposed;

    /// <summary>
    ///     Initializes a new instance of the <see cref="MmapTransport" /> class.
    /// </summary>
    /// <param name="logger">Optional logger instance.</param>
    /// <param name="maxDataSize">Maximum data size in bytes. Defaults to 100 MB.</param>
    /// <param name="tempDirectory">
    ///     Base directory for file-backed mmap (macOS only).
    ///     Defaults to system temp directory.
    /// </param>
    public MmapTransport(
        ILogger<MmapTransport>? logger = null,
        long maxDataSize = DefaultMaxDataSize,
        string? tempDirectory = null)
    {
        _logger = logger;
        _maxDataSize = maxDataSize;
        _cleanupTimer = new Timer(ProcessPendingCleanup, null, CleanupDelayMs, CleanupDelayMs);

        if (CurrentStrategy == MmapStrategy.FileBacked)
        {
            var baseDir = tempDirectory ?? Path.GetTempPath();
            _backingFileDirectory = Path.Combine(baseDir, $"aspose_mmap_{Environment.ProcessId}");
            Directory.CreateDirectory(_backingFileDirectory);
            CleanupOrphanedDirectories(baseDir);
        }

        _logger?.LogDebug("MmapTransport initialized with strategy: {Strategy}", CurrentStrategy);
    }

    /// <summary>
    ///     Gets the count of active memory-mapped files.
    ///     Used for diagnostics and monitoring.
    /// </summary>
    public int ActiveMmapCount => _activeMmaps.Count;

    /// <summary>
    ///     Gets the count of mmaps pending delayed cleanup.
    /// </summary>
    public int PendingCleanupCount => _pendingCleanup.Count;

    /// <summary>
    ///     Disposes all active memory-mapped files and releases resources.
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <inheritdoc />
    public string Mode => "mmap";

    /// <summary>
    ///     Sends data to the extension process using memory-mapped files.
    /// </summary>
    /// <param name="process">The extension process to send data to.</param>
    /// <param name="data">The binary data to send.</param>
    /// <param name="metadata">Metadata about the snapshot being sent.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>
    ///     <c>true</c> if the data was sent successfully; otherwise, <c>false</c>.
    /// </returns>
    /// <exception cref="ObjectDisposedException">Thrown when this instance has been disposed.</exception>
    public async Task<bool> SendAsync(
        Process process,
        byte[] data,
        ExtensionMetadata metadata,
        CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (process.HasExited)
            return false;

        if (data.Length > _maxDataSize)
        {
            _logger?.LogWarning(
                "Data size ({Size} bytes) exceeds maximum allowed size ({MaxSize} bytes) for session {SessionId}",
                data.Length, _maxDataSize, metadata.SessionId);
            return false;
        }

        if (_activeMmaps.Count >= MaxActiveMmaps)
            EvictOldestMmaps();

        MmapEntry? entry = null;
        string? mmapName = null;

        try
        {
            mmapName = GenerateMmapName(metadata.SessionId, metadata.SequenceNumber);
            entry = CreateMmapEntry(mmapName, data.Length);

            using (var accessor = entry.Value.File.CreateViewAccessor(0, data.Length, MemoryMappedFileAccess.Write))
            {
                accessor.WriteArray(0, data, 0, data.Length);
            }

            _activeMmaps[mmapName] = entry.Value;

            metadata.MmapName = mmapName;
            metadata.TransportMode = Mode;
            metadata.DataSize = data.Length;
            metadata.Checksum = Crc32.HashToUInt32(data);

            if (entry.Value.IsFileBacked)
                metadata.FilePath = entry.Value.BackingFilePath;

            var json = JsonSerializer.Serialize(metadata);
            await process.StandardInput.WriteLineAsync(json.AsMemory(), cancellationToken);
            await process.StandardInput.FlushAsync(cancellationToken);

            return true;
        }
        catch (Exception ex)
        {
            if (mmapName != null)
                _activeMmaps.TryRemove(mmapName, out _);

            CleanupEntry(entry);

            _logger?.LogWarning(ex,
                "Failed to send snapshot via mmap transport for session {SessionId}",
                metadata.SessionId);
            return false;
        }
    }

    /// <summary>
    ///     Cleans up resources associated with a snapshot.
    ///     Uses delayed cleanup to give child process time to finish reading.
    /// </summary>
    /// <param name="metadata">The metadata of the snapshot to clean up.</param>
    public void Cleanup(ExtensionMetadata metadata)
    {
        if (string.IsNullOrEmpty(metadata.MmapName))
            return;

        if (_activeMmaps.TryRemove(metadata.MmapName, out var entry))
        {
            if (_pendingCleanup.Count >= MaxPendingCleanup) DrainExcessPendingCleanup();

            _pendingCleanup.Enqueue((metadata.MmapName, entry, DateTime.UtcNow));
            _logger?.LogDebug(
                "Scheduled delayed cleanup for mmap: {MmapName}",
                metadata.MmapName);
        }
    }

    /// <summary>
    ///     Forces immediate cleanup of an mmap, bypassing delayed cleanup.
    ///     Use this when the extension has crashed and we know it won't read the data.
    /// </summary>
    /// <param name="mmapName">The name of the mmap to clean up.</param>
    /// <returns>True if the mmap was found and cleaned up.</returns>
    public bool ForceCleanup(string mmapName)
    {
        if (string.IsNullOrEmpty(mmapName))
            return false;

        if (_activeMmaps.TryRemove(mmapName, out var entry))
            try
            {
                CleanupEntry(entry);
                _logger?.LogDebug("Force cleaned up mmap: {MmapName}", mmapName);
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "Failed to force cleanup mmap: {MmapName}", mmapName);
            }

        return false;
    }

    /// <summary>
    ///     Determines the appropriate mmap strategy for the current platform.
    /// </summary>
    /// <returns>The mmap strategy to use.</returns>
    private static MmapStrategy DetermineStrategy()
    {
        if (OperatingSystem.IsWindows())
            return MmapStrategy.WindowsNamed;
        if (OperatingSystem.IsLinux())
            return MmapStrategy.LinuxPosix;
        return MmapStrategy.FileBacked;
    }

    /// <summary>
    ///     Creates a memory-mapped file entry using the appropriate platform strategy.
    /// </summary>
    /// <param name="mmapName">The name for the memory-mapped file.</param>
    /// <param name="dataLength">The size of data to store.</param>
    /// <returns>A new MmapEntry.</returns>
    /// <exception cref="IOException">Thrown when mmap creation fails.</exception>
    /// <exception cref="PlatformNotSupportedException">Thrown when platform strategy is unknown.</exception>
    private MmapEntry CreateMmapEntry(string mmapName, int dataLength)
    {
        return CurrentStrategy switch
        {
            MmapStrategy.WindowsNamed => CreateNamedMmap(mmapName, dataLength),
            MmapStrategy.LinuxPosix => CreateNamedMmap(mmapName, dataLength),
            MmapStrategy.FileBacked => CreateFileBackedMmap(mmapName, dataLength),
            _ => throw new PlatformNotSupportedException($"Unsupported mmap strategy: {CurrentStrategy}")
        };
    }

    /// <summary>
    ///     Creates a named memory-mapped file (Windows/Linux).
    /// </summary>
    /// <param name="mmapName">The name for the memory-mapped file.</param>
    /// <param name="dataLength">The size of data to store.</param>
    /// <returns>A new MmapEntry with memory-only storage.</returns>
    private static MmapEntry CreateNamedMmap(string mmapName, int dataLength)
    {
        var mmf = MemoryMappedFile.CreateNew(mmapName, dataLength);
        return new MmapEntry(mmf);
    }

    /// <summary>
    ///     Creates a file-backed memory-mapped file (macOS and other platforms).
    /// </summary>
    /// <param name="mmapName">The name for the memory-mapped file.</param>
    /// <param name="dataLength">The size of data to store.</param>
    /// <returns>A new MmapEntry with file-backed storage.</returns>
    /// <exception cref="InvalidOperationException">Thrown when backing directory is not initialized.</exception>
    private MmapEntry CreateFileBackedMmap(string mmapName, int dataLength)
    {
        if (string.IsNullOrEmpty(_backingFileDirectory))
            throw new InvalidOperationException("Backing file directory not initialized");

        var fileName = mmapName.Replace("/", "_").Replace("\\", "_");
        var filePath = Path.Combine(_backingFileDirectory, $"{fileName}.mmap");

        FileStream? fileStream = null;
        MemoryMappedFile? mmf = null;

        try
        {
            fileStream = new FileStream(
                filePath,
                FileMode.Create,
                FileAccess.ReadWrite,
                FileShare.ReadWrite,
                4096,
                FileOptions.None);

            fileStream.SetLength(dataLength);

            mmf = MemoryMappedFile.CreateFromFile(
                fileStream,
                null,
                dataLength,
                MemoryMappedFileAccess.ReadWrite,
                HandleInheritability.None,
                true);

            return new MmapEntry(mmf, fileStream, filePath);
        }
        catch
        {
            mmf?.Dispose();
            fileStream?.Dispose();
            CleanupBackingFile(filePath);
            throw;
        }
    }

    /// <summary>
    ///     Safely deletes a backing file, ignoring errors.
    /// </summary>
    /// <param name="filePath">The path of the file to delete.</param>
    private static void CleanupBackingFile(string? filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return;

        try
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
        catch
        {
            // Ignore file deletion errors during cleanup
        }
    }

    /// <summary>
    ///     Cleans up all resources associated with an mmap entry.
    /// </summary>
    /// <param name="entry">The entry to clean up, or null.</param>
    private void CleanupEntry(MmapEntry? entry)
    {
        if (entry == null)
            return;

        var e = entry.Value;

        try
        {
            e.File.Dispose();
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Failed to dispose MemoryMappedFile");
        }

        if (e.BackingStream != null)
            try
            {
                e.BackingStream.Dispose();
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "Failed to dispose backing FileStream");
            }

        CleanupBackingFile(e.BackingFilePath);
    }

    /// <summary>
    ///     Cleans up all resources associated with an mmap entry (non-nullable overload).
    /// </summary>
    /// <param name="entry">The entry to clean up.</param>
    private void CleanupEntry(MmapEntry entry)
    {
        try
        {
            entry.File.Dispose();
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Failed to dispose MemoryMappedFile");
        }

        if (entry.BackingStream != null)
            try
            {
                entry.BackingStream.Dispose();
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "Failed to dispose backing FileStream");
            }

        CleanupBackingFile(entry.BackingFilePath);
    }

    /// <summary>
    ///     Processes the pending cleanup queue, disposing mmaps that have waited long enough.
    /// </summary>
    /// <param name="state">Timer state (not used).</param>
    private void ProcessPendingCleanup(object? state)
    {
        if (_disposed)
            return;

        var now = DateTime.UtcNow;
        var delayThreshold = TimeSpan.FromMilliseconds(CleanupDelayMs);

        while (_pendingCleanup.TryPeek(out var item) && now - item.ScheduledAt >= delayThreshold)
        {
            if (!_pendingCleanup.TryDequeue(out item))
                break;

            try
            {
                CleanupEntry(item.Entry);
                _logger?.LogDebug("Disposed mmap after delay: {MmapName}", item.Name);
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "Failed to dispose mmap during delayed cleanup: {MmapName}", item.Name);
            }
        }
    }

    /// <summary>
    ///     Drains excess pending cleanup entries when the queue limit is reached.
    ///     Force-disposes items immediately to prevent unbounded growth.
    /// </summary>
    private void DrainExcessPendingCleanup()
    {
        var drainCount = 0;
        var targetCount = MaxPendingCleanup / 2;

        while (_pendingCleanup.Count > targetCount && _pendingCleanup.TryDequeue(out var item))
            try
            {
                CleanupEntry(item.Entry);
                drainCount++;
            }
            catch (Exception ex)
            {
                _logger?.LogDebug(ex, "Failed to force dispose mmap during queue drain: {MmapName}", item.Name);
            }

        if (drainCount > 0)
            _logger?.LogWarning(
                "Pending cleanup queue exceeded limit ({Limit}), force-disposed {Count} mmap(s). " +
                "Extensions may not be acknowledging snapshots promptly.",
                MaxPendingCleanup, drainCount);
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

        _disposed = true;

        if (disposing)
        {
            _cleanupTimer.Dispose();

            foreach (var kvp in _activeMmaps)
                CleanupEntry(kvp.Value);

            _activeMmaps.Clear();

            while (_pendingCleanup.TryDequeue(out var item))
                CleanupEntry(item.Entry);

            if (!string.IsNullOrEmpty(_backingFileDirectory))
                try
                {
                    if (Directory.Exists(_backingFileDirectory))
                        Directory.Delete(_backingFileDirectory, true);
                }
                catch (Exception ex)
                {
                    _logger?.LogDebug(ex,
                        "Failed to cleanup mmap backing directory: {Directory}",
                        _backingFileDirectory);
                }
        }
    }

    /// <summary>
    ///     Generates a unique memory-mapped file name for a snapshot.
    ///     Includes ProcessId and counter to prevent collisions.
    ///     For Linux, prepends "/" to comply with POSIX shm_open naming requirements.
    /// </summary>
    /// <param name="sessionId">The session identifier.</param>
    /// <param name="sequenceNumber">The sequence number of the snapshot.</param>
    /// <returns>A unique name for the memory-mapped file.</returns>
    private static string GenerateMmapName(string sessionId, long sequenceNumber)
    {
        var counter = Interlocked.Increment(ref _mmapCounter);
        var shortSessionId = sessionId.Length > 20 ? sessionId[..20] : sessionId;
        var baseName = $"aspose_{Environment.ProcessId}_{counter}_{shortSessionId}_{sequenceNumber}";

        return CurrentStrategy switch
        {
            MmapStrategy.LinuxPosix => "/" + baseName,
            _ => baseName
        };
    }

    /// <summary>
    ///     Evicts oldest mmaps to make room for new ones.
    ///     Uses delayed cleanup to give extensions a final chance to read data.
    ///     Called when the mmap limit is reached.
    /// </summary>
    /// <remarks>
    ///     <para>Thread safety considerations:</para>
    ///     <list type="bullet">
    ///         <item>Uses lock to prevent concurrent eviction attempts</item>
    ///         <item>Re-checks count after acquiring lock (another thread may have evicted)</item>
    ///         <item>Uses ConcurrentDictionary.TryRemove for atomic removal</item>
    ///     </list>
    ///     <para>
    ///         There is a potential race between the count check at the call site
    ///         and the eviction here. Multiple threads may pass the count check
    ///         simultaneously, but the lock ensures only one evicts at a time.
    ///         This is acceptable as slightly exceeding the limit temporarily
    ///         is preferable to the overhead of locking on every send.
    ///     </para>
    /// </remarks>
    private void EvictOldestMmaps()
    {
        lock (_evictionLock)
        {
            if (_activeMmaps.Count < MaxActiveMmaps)
                return;

            var countToEvict = (int)(MaxActiveMmaps * EvictionPercentage);
            countToEvict = Math.Max(countToEvict, 1);

            var toEvict = _activeMmaps
                .OrderBy(kvp => kvp.Value.CreatedAt)
                .Take(countToEvict)
                .Select(kvp => kvp.Key)
                .ToList();

            foreach (var key in toEvict)
                if (_activeMmaps.TryRemove(key, out var entry))
                    _pendingCleanup.Enqueue((key, entry, DateTime.UtcNow));

            _logger?.LogWarning(
                "Evicted {Count} oldest mmap(s) due to limit ({Limit}). " +
                "Data will be disposed after {DelayMs}ms delay. " +
                "Extensions may not be acknowledging snapshots promptly.",
                toEvict.Count, MaxActiveMmaps, CleanupDelayMs);
        }
    }

    /// <summary>
    ///     Cleans up orphaned mmap directories from previous crashed processes.
    /// </summary>
    /// <param name="baseDirectory">The base directory to scan.</param>
    private void CleanupOrphanedDirectories(string baseDirectory)
    {
        try
        {
            var orphanDirs = Directory.GetDirectories(baseDirectory, "aspose_mmap_*");
            var cleanedCount = 0;

            foreach (var dir in orphanDirs)
            {
                if (dir == _backingFileDirectory)
                    continue;

                var dirName = Path.GetFileName(dir);
                if (string.IsNullOrEmpty(dirName) || !dirName.StartsWith("aspose_mmap_"))
                    continue;

                var pidPart = dirName["aspose_mmap_".Length..];
                if (!int.TryParse(pidPart, out var pid))
                    continue;

                if (IsOrphanedDirectory(dir, pid))
                    try
                    {
                        Directory.Delete(dir, true);
                        cleanedCount++;
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogDebug(ex, "Failed to delete orphaned mmap directory: {Directory}", dir);
                    }
            }

            if (cleanedCount > 0)
                _logger?.LogInformation(
                    "Cleaned up {Count} orphaned mmap director(ies) from previous crashed runs",
                    cleanedCount);
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to cleanup orphaned mmap directories");
        }
    }

    /// <summary>
    ///     Determines if a directory is orphaned (process no longer exists or was restarted).
    /// </summary>
    /// <param name="directoryPath">The path to the directory.</param>
    /// <param name="pid">The process ID extracted from the directory name.</param>
    /// <returns>True if the directory is orphaned and can be deleted.</returns>
    private static bool IsOrphanedDirectory(string directoryPath, int pid)
    {
        try
        {
            using var process = Process.GetProcessById(pid);

            if (process.HasExited)
                return true;

            var dirCreationTime = Directory.GetCreationTimeUtc(directoryPath);
            var processStartTime = process.StartTime.ToUniversalTime();

            return dirCreationTime < processStartTime;
        }
        catch (ArgumentException)
        {
            return true;
        }
        catch (InvalidOperationException)
        {
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    ///     Mmap creation strategy based on platform capabilities.
    /// </summary>
    private enum MmapStrategy
    {
        /// <summary>Windows: Named shared memory using kernel objects.</summary>
        WindowsNamed,

        /// <summary>Linux: POSIX shm_open via /dev/shm/.</summary>
        LinuxPosix,

        /// <summary>macOS/Other: File-backed mmap (named mmap not supported).</summary>
        FileBacked
    }

    /// <summary>
    ///     Entry tracking a memory-mapped file and its associated resources.
    /// </summary>
    private readonly struct MmapEntry
    {
        /// <summary>
        ///     The memory-mapped file instance.
        /// </summary>
        public MemoryMappedFile File { get; }

        /// <summary>
        ///     UTC time when this mmap was created.
        /// </summary>
        public DateTime CreatedAt { get; }

        /// <summary>
        ///     Backing file stream for file-backed mmap (macOS only).
        /// </summary>
        public FileStream? BackingStream { get; }

        /// <summary>
        ///     Path to the backing file for cleanup (macOS only).
        /// </summary>
        public string? BackingFilePath { get; }

        /// <summary>
        ///     Whether this entry uses file-backed storage.
        /// </summary>
        public bool IsFileBacked => BackingStream != null;

        /// <summary>
        ///     Creates a memory-only mmap entry (Windows/Linux).
        /// </summary>
        /// <param name="file">The memory-mapped file instance.</param>
        public MmapEntry(MemoryMappedFile file)
        {
            File = file;
            CreatedAt = DateTime.UtcNow;
            BackingStream = null;
            BackingFilePath = null;
        }

        /// <summary>
        ///     Creates a file-backed mmap entry (macOS).
        /// </summary>
        /// <param name="file">The memory-mapped file instance.</param>
        /// <param name="backingStream">The backing file stream.</param>
        /// <param name="backingFilePath">The path to the backing file.</param>
        public MmapEntry(MemoryMappedFile file, FileStream backingStream, string backingFilePath)
        {
            File = file;
            CreatedAt = DateTime.UtcNow;
            BackingStream = backingStream;
            BackingFilePath = backingFilePath;
        }
    }
}
