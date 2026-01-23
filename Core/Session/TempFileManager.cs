using System.Text.Json;

namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Manages temporary files created by session disconnect behavior.
///     Handles cleanup of expired files and recovery of saved sessions.
/// </summary>
public class TempFileManager : IHostedService, IDisposable // NOSONAR S3881 - Simple dispose pattern sufficient
{
    private const string TempFilePrefix = "aspose_session_";
    private const string MetadataExtension = ".meta.json";
    private readonly SessionConfig _config;

    private readonly ILogger<TempFileManager>? _logger;
    private Timer? _cleanupTimer;
    private int _disposed;

    /// <summary>
    ///     Creates a new temp file manager
    /// </summary>
    /// <param name="config">Session configuration</param>
    /// <param name="loggerFactory">Logger factory for logging</param>
    public TempFileManager(SessionConfig config, ILoggerFactory? loggerFactory = null)
    {
        _config = config;
        _logger = loggerFactory?.CreateLogger<TempFileManager>();
    }

    /// <summary>
    ///     Disposes the temp file manager.
    ///     Thread-safe: uses Interlocked to prevent double-dispose.
    /// </summary>
    public void Dispose()
    {
        // Atomically set _disposed to 1, return previous value
        // If previous value was already 1, another thread already disposed
        if (Interlocked.Exchange(ref _disposed, 1) == 1)
            return;
        _cleanupTimer?.Dispose();
    }

    /// <summary>
    ///     Starts the temp file manager service.
    ///     Performs initial cleanup and starts periodic cleanup timer.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>A completed task</returns>
    public Task StartAsync(CancellationToken cancellationToken)
    {
        if (!_config.Enabled)
        {
            _logger?.LogDebug("Session management disabled, skipping temp file manager");
            return Task.CompletedTask;
        }

        _logger?.LogInformation("Starting temp file manager (retention: {Hours} hours)", _config.TempRetentionHours);

        var cleanupResult = CleanupExpiredFiles();
        _logger?.LogInformation("Startup cleanup completed: {Deleted} files deleted, {Errors} errors",
            cleanupResult.DeletedCount, cleanupResult.ErrorCount);

        _cleanupTimer = new Timer(
            PeriodicCleanup,
            null,
            TimeSpan.FromHours(1),
            TimeSpan.FromHours(1));

        return Task.CompletedTask;
    }

    /// <summary>
    ///     Stops the temp file manager service.
    ///     Stops the periodic cleanup timer.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>A completed task</returns>
    public Task StopAsync(CancellationToken cancellationToken)
    {
        _logger?.LogInformation("Stopping temp file manager");
        _cleanupTimer?.Change(Timeout.Infinite, 0);
        return Task.CompletedTask;
    }

    /// <summary>
    ///     Checks if the requestor can access a temp file based on isolation mode
    /// </summary>
    /// <param name="requestor">The identity of the requestor</param>
    /// <param name="metadata">The temp file metadata containing owner info</param>
    /// <returns>True if access is allowed</returns>
    private bool CanAccessTempFile(SessionIdentity requestor, TempFileMetadata metadata)
    {
        var owner = new SessionIdentity
        {
            GroupId = metadata.OwnerGroupId,
            UserId = metadata.OwnerUserId
        };
        return requestor.CanAccess(owner, _config.IsolationMode);
    }

    /// <summary>
    ///     Cleans up expired temporary files based on TempRetentionHours
    /// </summary>
    /// <returns>Cleanup result with statistics</returns>
    public CleanupResult CleanupExpiredFiles()
    {
        var result = new CleanupResult();
        var cutoffTime = DateTime.UtcNow.AddHours(-_config.TempRetentionHours);

        try
        {
            if (!Directory.Exists(_config.TempDirectory))
            {
                _logger?.LogDebug("Temp directory does not exist: {TempDir}", _config.TempDirectory);
                return result;
            }

            var metadataFiles = Directory.GetFiles(_config.TempDirectory, $"{TempFilePrefix}*{MetadataExtension}");

            foreach (var metadataPath in metadataFiles)
                try
                {
                    result.ScannedCount++;
                    var metadata = ReadMetadata(metadataPath);

                    if (metadata == null)
                    {
                        DeleteTempFileSet(metadataPath);
                        result.DeletedCount++;
                        continue;
                    }

                    if (metadata.SavedAt < cutoffTime)
                    {
                        DeleteTempFileSet(metadataPath);
                        result.DeletedCount++;
                        _logger?.LogDebug("Deleted expired temp file: {Path} (saved at {SavedAt})",
                            metadata.TempPath, metadata.SavedAt);
                    }
                }
                catch (Exception ex)
                {
                    result.ErrorCount++;
                    _logger?.LogWarning(ex, "Error processing temp file: {Path}", metadataPath);
                }

            var orphanedFiles = Directory.GetFiles(_config.TempDirectory, $"{TempFilePrefix}*")
                .Where(f => !f.EndsWith(MetadataExtension))
                .Where(f => !File.Exists(f + MetadataExtension));

            foreach (var orphanedFile in orphanedFiles)
                try
                {
                    var fileInfo = new FileInfo(orphanedFile);
                    if (fileInfo.LastWriteTimeUtc < cutoffTime)
                    {
                        File.Delete(orphanedFile);
                        result.DeletedCount++;
                        _logger?.LogDebug("Deleted orphaned temp file: {Path}", orphanedFile);
                    }
                }
                catch (Exception ex)
                {
                    result.ErrorCount++;
                    _logger?.LogWarning(ex, "Error deleting orphaned file: {Path}", orphanedFile);
                }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error during temp file cleanup");
            result.ErrorCount++;
        }

        return result;
    }

    /// <summary>
    ///     Lists all recoverable temporary files (no authorization check - returns all)
    /// </summary>
    /// <returns>List of recoverable file information</returns>
    public IEnumerable<RecoverableFileInfo> ListRecoverableFiles()
    {
        return ListRecoverableFiles(SessionIdentity.GetAnonymous());
    }

    /// <summary>
    ///     Lists recoverable temporary files visible to the requestor
    /// </summary>
    /// <param name="requestor">Requestor identity for filtering</param>
    /// <returns>List of recoverable file information</returns>
    public IEnumerable<RecoverableFileInfo> ListRecoverableFiles(SessionIdentity requestor)
    {
        var results = new List<RecoverableFileInfo>();

        try
        {
            if (!Directory.Exists(_config.TempDirectory))
                return results;

            var metadataFiles = Directory.GetFiles(_config.TempDirectory, $"{TempFilePrefix}*{MetadataExtension}");

            foreach (var metadataPath in metadataFiles)
                try
                {
                    var metadata = ReadMetadata(metadataPath);
                    if (metadata == null) continue;

                    if (!CanAccessTempFile(requestor, metadata))
                    {
                        _logger?.LogDebug("Access denied: {Requestor} cannot access temp file {SessionId}",
                            requestor, metadata.SessionId);
                        continue;
                    }

                    if (!File.Exists(metadata.TempPath)) continue;

                    var fileInfo = new FileInfo(metadata.TempPath);
                    var expiresAt = metadata.SavedAt.AddHours(_config.TempRetentionHours);

                    results.Add(new RecoverableFileInfo
                    {
                        SessionId = metadata.SessionId,
                        OriginalPath = metadata.OriginalPath,
                        TempPath = metadata.TempPath,
                        DocumentType = metadata.DocumentType,
                        SavedAt = metadata.SavedAt,
                        ExpiresAt = expiresAt,
                        FileSizeBytes = fileInfo.Length,
                        PromptOnReconnect = metadata.PromptOnReconnect,
                        OwnerGroupId = metadata.OwnerGroupId,
                        OwnerUserId = metadata.OwnerUserId
                    });
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "Error reading metadata: {Path}", metadataPath);
                }
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error listing recoverable files");
        }

        return results.OrderByDescending(r => r.SavedAt);
    }

    /// <summary>
    ///     Recovers a temporary file to the specified path (no authorization check)
    /// </summary>
    /// <param name="sessionId">Session ID to recover</param>
    /// <param name="targetPath">Target path (null = original path)</param>
    /// <param name="deleteAfterRecover">Whether to delete temp file after recovery</param>
    /// <returns>Recovery result</returns>
    public RecoverResult RecoverSession(string sessionId, string? targetPath = null, bool deleteAfterRecover = true)
    {
        return RecoverSession(sessionId, SessionIdentity.GetAnonymous(), targetPath, deleteAfterRecover);
    }

    /// <summary>
    ///     Recovers a temporary file to the specified path with authorization check
    /// </summary>
    /// <param name="sessionId">Session ID to recover</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    /// <param name="targetPath">Target path (null = original path)</param>
    /// <param name="deleteAfterRecover">Whether to delete temp file after recovery</param>
    /// <returns>Recovery result</returns>
    public RecoverResult RecoverSession(string sessionId, SessionIdentity requestor, string? targetPath = null,
        bool deleteAfterRecover = true)
    {
        var result = new RecoverResult { SessionId = sessionId };

        try
        {
            var metadataFiles =
                Directory.GetFiles(_config.TempDirectory, $"{TempFilePrefix}{sessionId}*{MetadataExtension}");

            if (metadataFiles.Length == 0)
            {
                result.Success = false;
                result.ErrorMessage = $"No recoverable session found: {sessionId}";
                return result;
            }

            var metadataPath = metadataFiles
                .Select(f => new { Path = f, Info = new FileInfo(f) })
                .OrderByDescending(x => x.Info.LastWriteTimeUtc)
                .First().Path;

            var metadata = ReadMetadata(metadataPath);
            if (metadata == null)
            {
                result.Success = false;
                result.ErrorMessage = "Failed to read session metadata";
                return result;
            }

            if (!CanAccessTempFile(requestor, metadata))
            {
                _logger?.LogWarning(
                    "Access denied: {Requestor} attempted to recover temp file {SessionId}",
                    requestor, sessionId);
                result.Success = false;
                result.ErrorMessage = $"No recoverable session found: {sessionId}";
                return result;
            }

            if (!File.Exists(metadata.TempPath))
            {
                result.Success = false;
                result.ErrorMessage = $"Temp file not found: {metadata.TempPath}";
                return result;
            }

            var destination = targetPath ?? metadata.OriginalPath;

            var targetDir = Path.GetDirectoryName(destination);
            if (!string.IsNullOrEmpty(targetDir) && !Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);

            File.Copy(metadata.TempPath, destination, true);

            result.Success = true;
            result.RecoveredPath = destination;
            result.OriginalPath = metadata.OriginalPath;

            _logger?.LogInformation("Recovered session {SessionId} to {Path}", sessionId, destination);

            if (deleteAfterRecover)
            {
                DeleteTempFileSet(metadataPath);
                _logger?.LogDebug("Deleted temp files after recovery: {SessionId}", sessionId);
            }
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = ex.Message;
            _logger?.LogError(ex, "Error recovering session {SessionId}", sessionId);
        }

        return result;
    }

    /// <summary>
    ///     Deletes a specific temporary session file (no authorization check)
    /// </summary>
    /// <param name="sessionId">Session ID to delete</param>
    /// <returns>True if deleted successfully</returns>
    public bool DeleteTempSession(string sessionId)
    {
        return DeleteTempSession(sessionId, SessionIdentity.GetAnonymous());
    }

    /// <summary>
    ///     Deletes a specific temporary session file with authorization check
    /// </summary>
    /// <param name="sessionId">Session ID to delete</param>
    /// <param name="requestor">Requestor identity for authorization</param>
    /// <returns>True if deleted successfully</returns>
    public bool DeleteTempSession(string sessionId, SessionIdentity requestor)
    {
        try
        {
            var metadataFiles =
                Directory.GetFiles(_config.TempDirectory, $"{TempFilePrefix}{sessionId}*{MetadataExtension}");

            if (metadataFiles.Length == 0)
                return false;

            var deleted = false;
            foreach (var metadataPath in metadataFiles)
            {
                var metadata = ReadMetadata(metadataPath);
                if (metadata == null)
                {
                    // Invalid metadata, allow cleanup
                    DeleteTempFileSet(metadataPath);
                    deleted = true;
                    continue;
                }

                if (!CanAccessTempFile(requestor, metadata))
                {
                    _logger?.LogWarning(
                        "Access denied: {Requestor} attempted to delete temp file {SessionId}",
                        requestor, sessionId);
                    continue;
                }

                DeleteTempFileSet(metadataPath);
                deleted = true;
            }

            if (deleted)
                _logger?.LogInformation("Deleted temp session: {SessionId}", sessionId);

            return deleted;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error deleting temp session {SessionId}", sessionId);
            return false;
        }
    }

    /// <summary>
    ///     Gets cleanup statistics
    /// </summary>
    /// <returns>Current temp file statistics</returns>
    public TempFileStats GetStats()
    {
        var stats = new TempFileStats();

        try
        {
            if (!Directory.Exists(_config.TempDirectory))
                return stats;

            var metadataFiles = Directory.GetFiles(_config.TempDirectory, $"{TempFilePrefix}*{MetadataExtension}");
            var cutoffTime = DateTime.UtcNow.AddHours(-_config.TempRetentionHours);

            foreach (var metadataPath in metadataFiles)
                try
                {
                    var metadata = ReadMetadata(metadataPath);
                    if (metadata == null) continue;

                    if (File.Exists(metadata.TempPath))
                    {
                        var fileInfo = new FileInfo(metadata.TempPath);
                        stats.TotalCount++;
                        stats.TotalSizeBytes += fileInfo.Length;

                        if (metadata.SavedAt < cutoffTime)
                            stats.ExpiredCount++;
                    }
                }
                catch
                {
                    // Ignore errors in stats collection
                }
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Error collecting temp file stats");
        }

        return stats;
    }

    /// <summary>
    ///     Periodic cleanup callback invoked by the timer
    /// </summary>
    /// <param name="state">Timer state (unused)</param>
    private void PeriodicCleanup(object? state)
    {
        try
        {
            var result = CleanupExpiredFiles();
            if (result.DeletedCount > 0 || result.ErrorCount > 0)
                _logger?.LogInformation("Periodic cleanup: {Deleted} deleted, {Errors} errors",
                    result.DeletedCount, result.ErrorCount);
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error in periodic cleanup");
        }
    }

    /// <summary>
    ///     Reads metadata from a metadata file
    /// </summary>
    /// <param name="metadataPath">Path to the metadata file</param>
    /// <returns>Deserialized metadata, or null if reading fails</returns>
    private static TempFileMetadata? ReadMetadata(string metadataPath)
    {
        try
        {
            var json = File.ReadAllText(metadataPath);
            return JsonSerializer.Deserialize<TempFileMetadata>(json);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    ///     Deletes a temp file and its metadata
    /// </summary>
    /// <param name="metadataPath">Path to the metadata file</param>
    private void DeleteTempFileSet(string metadataPath)
    {
        try
        {
            var metadata = ReadMetadata(metadataPath);

            if (metadata?.TempPath != null && File.Exists(metadata.TempPath))
                File.Delete(metadata.TempPath);

            if (File.Exists(metadataPath))
                File.Delete(metadataPath);
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Error deleting temp file set: {Path}", metadataPath);
        }
    }
}
