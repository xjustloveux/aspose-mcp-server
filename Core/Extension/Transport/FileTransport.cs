using System.Diagnostics;
using System.IO.Hashing;
using System.Text.Json;

namespace AsposeMcpServer.Core.Extension.Transport;

/// <summary>
///     File-based transport for extensions.
///     Writes data to a temporary file and sends the file path via stdin.
/// </summary>
// ReSharper disable once ClassWithVirtualMembersNeverInherited.Global - Virtual Dispose(bool) for proper IDisposable pattern
public class FileTransport : IExtensionTransport, IDisposable
{
    /// <summary>
    ///     Prefix for extension snapshot directories.
    /// </summary>
    private const string DirectoryPrefix = "ext_snapshots_";

    /// <summary>
    ///     Default maximum snapshot file size in bytes (100 MB).
    /// </summary>
    private const long DefaultMaxSnapshotSize = 100 * 1024 * 1024;

    /// <summary>
    ///     Default minimum free disk space in bytes (500 MB).
    /// </summary>
    private const long DefaultMinFreeDiskSpace = 500 * 1024 * 1024;

    /// <summary>
    ///     Counter for generating unique file names to prevent concurrent write conflicts.
    /// </summary>
    private static long _fileCounter;

    /// <summary>
    ///     Logger instance for diagnostic output.
    /// </summary>
    private readonly ILogger<FileTransport>? _logger;

    /// <summary>
    ///     Maximum snapshot file size in bytes.
    /// </summary>
    private readonly long _maxSnapshotSize;

    /// <summary>
    ///     Minimum free disk space in bytes required before writing.
    /// </summary>
    private readonly long _minFreeDiskSpace;

    /// <summary>
    ///     The directory path where temporary snapshot files are stored.
    /// </summary>
    private readonly string _tempDirectory;

    /// <summary>
    ///     Whether this instance has been disposed.
    /// </summary>
    private volatile bool _disposed;

    /// <summary>
    ///     Initializes a new instance of the <see cref="FileTransport" /> class.
    /// </summary>
    /// <param name="tempDirectory">Directory for temporary files.</param>
    /// <param name="logger">Optional logger instance.</param>
    /// <param name="maxSnapshotSize">Maximum snapshot size in bytes. Defaults to 100 MB.</param>
    /// <param name="minFreeDiskSpace">Minimum free disk space in bytes. Defaults to 500 MB.</param>
    /// <exception cref="ArgumentException">Thrown when tempDirectory is null or empty.</exception>
    public FileTransport(
        string tempDirectory,
        ILogger<FileTransport>? logger = null,
        long maxSnapshotSize = DefaultMaxSnapshotSize,
        long minFreeDiskSpace = DefaultMinFreeDiskSpace)
    {
        if (string.IsNullOrWhiteSpace(tempDirectory))
            throw new ArgumentException("Temp directory path cannot be null or empty", nameof(tempDirectory));

        _tempDirectory = tempDirectory;
        _logger = logger;
        _maxSnapshotSize = maxSnapshotSize;
        _minFreeDiskSpace = minFreeDiskSpace;
        CleanupOrphanedFiles();
        Directory.CreateDirectory(_tempDirectory);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <inheritdoc />
    public string Mode => "file";

    /// <inheritdoc />
    public async Task<bool> SendAsync(
        Process process,
        byte[] data,
        ExtensionMetadata metadata,
        CancellationToken cancellationToken = default)
    {
        if (process.HasExited)
            return false;

        if (data.Length > _maxSnapshotSize)
        {
            _logger?.LogWarning(
                "Snapshot size ({Size} bytes) exceeds maximum allowed size ({MaxSize} bytes) for session {SessionId}",
                data.Length, _maxSnapshotSize, metadata.SessionId);
            return false;
        }

        if (!HasSufficientDiskSpace(data.Length))
        {
            _logger?.LogWarning(
                "Insufficient disk space for snapshot ({Size} bytes). " +
                "Free space is below {MinFree} MB threshold for session {SessionId}",
                data.Length, _minFreeDiskSpace / (1024 * 1024), metadata.SessionId);
            return false;
        }

        if (!EnsureTempDirectoryExists())
        {
            _logger?.LogWarning(
                "Temp directory does not exist and could not be created for session {SessionId}",
                metadata.SessionId);
            return false;
        }

        string? filePath = null;
        try
        {
            var sanitizedSessionId = SanitizeFileName(metadata.SessionId);
            var sanitizedFormat = SanitizeFileName(metadata.OutputFormat);
            var uniqueId = Interlocked.Increment(ref _fileCounter);
            var fileName = $"ext_snapshot_{sanitizedSessionId}_{metadata.SequenceNumber}_{uniqueId}.{sanitizedFormat}";
            filePath = Path.Combine(_tempDirectory, fileName);

            await File.WriteAllBytesAsync(filePath, data, cancellationToken);

            metadata.FilePath = filePath;
            metadata.TransportMode = Mode;
            metadata.DataSize = data.Length;
            metadata.Checksum = Crc32.HashToUInt32(data);

            var json = JsonSerializer.Serialize(metadata);
            await process.StandardInput.WriteLineAsync(json.AsMemory(), cancellationToken);
            await process.StandardInput.FlushAsync(cancellationToken);

            return true;
        }
        catch (IOException ioEx) when (IsDiskSpaceError(ioEx))
        {
            CleanupFailedFile(filePath);
            _logger?.LogError(ioEx,
                "Disk space exhausted while writing snapshot for session {SessionId}. " +
                "Free up disk space or reduce snapshot size.",
                metadata.SessionId);
            return false;
        }
        catch (UnauthorizedAccessException uaEx)
        {
            CleanupFailedFile(filePath);
            _logger?.LogError(uaEx,
                "Permission denied writing snapshot for session {SessionId}. " +
                "Check write permissions for temp directory: {TempDir}",
                metadata.SessionId, _tempDirectory);
            return false;
        }
        catch (DirectoryNotFoundException)
        {
            CleanupFailedFile(filePath);
            _logger?.LogWarning(
                "Temp directory was deleted during snapshot write for session {SessionId}",
                metadata.SessionId);
            return false;
        }
        catch (Exception ex)
        {
            CleanupFailedFile(filePath);
            _logger?.LogWarning(ex,
                "Failed to send snapshot via file transport for session {SessionId}",
                metadata.SessionId);
            return false;
        }
    }

    /// <inheritdoc />
    public void Cleanup(ExtensionMetadata metadata)
    {
        if (string.IsNullOrEmpty(metadata.FilePath))
            return;

        try
        {
            if (File.Exists(metadata.FilePath))
                File.Delete(metadata.FilePath);
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex,
                "Failed to cleanup snapshot file: {FilePath}",
                metadata.FilePath);
        }
    }

    /// <summary>
    ///     Cleans up orphaned snapshot files from previous runs that may have crashed.
    ///     Removes any existing files in the temp directory before starting fresh.
    /// </summary>
    private void CleanupOrphanedFiles()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                var orphanedFiles = Directory.GetFiles(_tempDirectory, "ext_snapshot_*");
                if (orphanedFiles.Length > 0)
                {
                    _logger?.LogInformation(
                        "Cleaning up {Count} orphaned snapshot file(s) from previous run",
                        orphanedFiles.Length);

                    foreach (var file in orphanedFiles)
                        try
                        {
                            File.Delete(file);
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogDebug(ex,
                                "Failed to delete orphaned file: {FilePath}",
                                file);
                        }
                }
            }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex,
                "Failed to cleanup orphaned files in temp directory: {TempDirectory}",
                _tempDirectory);
        }
    }

    /// <summary>
    ///     Cleans up a partially written file after a failure.
    /// </summary>
    /// <param name="filePath">Path to the file to clean up.</param>
    /// <remarks>
    ///     This is a best-effort cleanup. Exceptions are intentionally suppressed
    ///     because this is called during error handling paths where we don't want
    ///     to mask the original error. The file may be locked by another process
    ///     or already deleted.
    /// </remarks>
    private static void CleanupFailedFile(string? filePath)
    {
        if (filePath == null)
            return;

        try
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
        // ReSharper disable once EmptyGeneralCatchClause
        catch
        {
        }
    }

    /// <summary>
    ///     Ensures the temp directory exists, creating it if necessary.
    /// </summary>
    /// <returns>True if the directory exists or was created successfully.</returns>
    private bool EnsureTempDirectoryExists()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
                return true;

            Directory.CreateDirectory(_tempDirectory);
            _logger?.LogInformation(
                "Recreated temp directory that was deleted: {TempDirectory}",
                _tempDirectory);
            return true;
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex,
                "Failed to ensure temp directory exists: {TempDirectory}",
                _tempDirectory);
            return false;
        }
    }

    /// <summary>
    ///     Checks if there is sufficient disk space for writing a file.
    /// </summary>
    /// <param name="requiredBytes">The number of bytes required.</param>
    /// <returns>True if there is sufficient disk space.</returns>
    private bool HasSufficientDiskSpace(long requiredBytes)
    {
        try
        {
            var driveInfo = new DriveInfo(Path.GetPathRoot(_tempDirectory) ?? _tempDirectory);
            var availableBytes = driveInfo.AvailableFreeSpace;
            return availableBytes >= requiredBytes + _minFreeDiskSpace;
        }
        catch (Exception ex)
        {
            _logger?.LogDebug(ex, "Failed to check disk space, allowing write attempt");
            return true;
        }
    }

    /// <summary>
    ///     Checks if an IOException is related to disk space exhaustion.
    /// </summary>
    /// <param name="ex">The exception to check.</param>
    /// <returns>True if the error is disk space related.</returns>
    private static bool IsDiskSpaceError(IOException ex)
    {
        var hResult = ex.HResult & 0xFFFF;
        return hResult == 112 || hResult == 39 || hResult == 28 ||
               ex.Message.Contains("disk", StringComparison.OrdinalIgnoreCase) ||
               ex.Message.Contains("space", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     Sanitizes a string for use in file names by removing invalid characters.
    /// </summary>
    /// <param name="input">The input string to sanitize.</param>
    /// <returns>A sanitized string safe for file names.</returns>
    /// <remarks>
    ///     <para>Edge cases handled:</para>
    ///     <list type="bullet">
    ///         <item>Null or empty input: returns "unknown"</item>
    ///         <item>Invalid characters (e.g., /, \, :, *, ?, ", &lt;, &gt;, |): replaced with underscore</item>
    ///         <item>Unicode characters: preserved if not in invalid set</item>
    ///     </list>
    ///     <para>
    ///         Note: This method does not truncate long names. File path length limits
    ///         are enforced at the OS level and will result in an IOException if exceeded.
    ///     </para>
    /// </remarks>
    private static string SanitizeFileName(string input)
    {
        if (string.IsNullOrEmpty(input))
            return "unknown";

        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = new char[input.Length];
        var index = 0;

        foreach (var c in input) sanitized[index++] = Array.IndexOf(invalidChars, c) >= 0 ? '_' : c;

        return new string(sanitized);
    }

    /// <summary>
    ///     Disposes resources and cleans up the temporary directory.
    /// </summary>
    /// <param name="disposing">
    ///     <c>true</c> to dispose managed resources; <c>false</c> to dispose only unmanaged resources.
    /// </param>
    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing) CleanupTempDirectory();

        _disposed = true;
    }

    /// <summary>
    ///     Cleans up the temporary directory and all files within it.
    /// </summary>
    private void CleanupTempDirectory()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, true);
                _logger?.LogDebug("Cleaned up temp directory: {TempDirectory}", _tempDirectory);
            }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex,
                "Failed to cleanup temp directory: {TempDirectory}",
                _tempDirectory);
        }
    }

    /// <summary>
    ///     Cleans up orphaned snapshot directories from previous process runs that crashed.
    ///     Removes directories whose process IDs no longer exist.
    /// </summary>
    /// <param name="baseTempDirectory">The base temporary directory containing snapshot subdirectories.</param>
    /// <param name="logger">Optional logger instance.</param>
    public static void CleanupOrphanedDirectories(string baseTempDirectory, ILogger? logger = null)
    {
        try
        {
            if (!Directory.Exists(baseTempDirectory))
                return;

            var orphanDirs = Directory.GetDirectories(baseTempDirectory, $"{DirectoryPrefix}*");
            var cleanedCount = 0;

            foreach (var dir in orphanDirs)
            {
                var dirName = Path.GetFileName(dir);
                var pidPart = dirName[DirectoryPrefix.Length..];

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
                        logger?.LogDebug(ex,
                            "Failed to delete orphaned directory: {Directory}",
                            dir);
                    }
            }

            if (cleanedCount > 0)
                logger?.LogInformation(
                    "Cleaned up {Count} orphaned snapshot director(ies) from previous crashed runs",
                    cleanedCount);
        }
        catch (Exception ex)
        {
            logger?.LogWarning(ex,
                "Failed to cleanup orphaned directories in: {TempDirectory}",
                baseTempDirectory);
        }
    }

    /// <summary>
    ///     Determines if a directory is orphaned by checking process status and creation time.
    ///     Handles PID reuse by comparing directory creation time with process start time.
    /// </summary>
    /// <param name="directoryPath">The directory path to check.</param>
    /// <param name="pid">The process ID from the directory name.</param>
    /// <returns>True if the directory is orphaned and safe to delete.</returns>
    private static bool IsOrphanedDirectory(string directoryPath, int pid)
    {
        try
        {
            using var process = Process.GetProcessById(pid);
            if (process.HasExited)
                return true;

            var dirCreationTime = Directory.GetCreationTimeUtc(directoryPath);
            var processStartTime = process.StartTime.ToUniversalTime();

            if (dirCreationTime < processStartTime)
                return true;

            return false;
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
    ///     Generates a directory name with the current process ID.
    /// </summary>
    /// <param name="baseTempDirectory">The base temporary directory.</param>
    /// <returns>A directory path with PID suffix.</returns>
    public static string GenerateDirectoryWithPid(string baseTempDirectory)
    {
        var pid = Environment.ProcessId;
        return Path.Combine(baseTempDirectory, $"{DirectoryPrefix}{pid}");
    }
}
