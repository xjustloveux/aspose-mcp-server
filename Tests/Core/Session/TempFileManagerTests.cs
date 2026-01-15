using System.Text.Json;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Tests.Core.Session;

/// <summary>
///     Unit tests for TempFileManager class
/// </summary>
public class TempFileManagerTests : IDisposable
{
    private readonly SessionConfig _config;
    private readonly TempFileManager _manager;
    private readonly string _tempDir;

    public TempFileManagerTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "TempFileManagerTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(_tempDir);

        _config = new SessionConfig
        {
            Enabled = true,
            TempDirectory = _tempDir,
            TempRetentionHours = 24
        };

        _manager = new TempFileManager(_config);
    }

    public void Dispose()
    {
        _manager.Dispose();

        try
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, true);
        }
        catch
        {
            // Ignore cleanup errors
        }
    }

    #region Helper Methods

    private (string docPath, string metaPath) CreateTempFile(string sessionId, DateTime savedAt,
        string? originalPath = null)
    {
        var timestamp = savedAt.ToString("yyyyMMddHHmmss");
        var docPath = Path.Combine(_tempDir, $"aspose_session_{sessionId}_{timestamp}.docx");
        var metaPath = docPath + ".meta.json";

        File.WriteAllText(docPath, "Test document content");

        var metadata = new
        {
            SessionId = sessionId,
            OriginalPath = originalPath ?? $"/test/{sessionId}.docx",
            TempPath = docPath,
            DocumentType = "Word",
            SavedAt = savedAt,
            PromptOnReconnect = false
        };
        File.WriteAllText(metaPath, JsonSerializer.Serialize(metadata));

        return (docPath, metaPath);
    }

    #endregion

    #region CleanupExpiredFiles Tests

    [Fact]
    public void CleanupExpiredFiles_EmptyDirectory_ShouldReturnZeroCounts()
    {
        var result = _manager.CleanupExpiredFiles();

        Assert.Equal(0, result.ScannedCount);
        Assert.Equal(0, result.DeletedCount);
        Assert.Equal(0, result.ErrorCount);
    }

    [Fact]
    public void CleanupExpiredFiles_NonExpiredFiles_ShouldNotDelete()
    {
        CreateTempFile("sess_test1234", DateTime.UtcNow);

        var result = _manager.CleanupExpiredFiles();

        Assert.Equal(1, result.ScannedCount);
        Assert.Equal(0, result.DeletedCount);
    }

    [Fact]
    public void CleanupExpiredFiles_ExpiredFiles_ShouldDelete()
    {
        // Create an expired temp file (older than retention hours)
        var expiredTime = DateTime.UtcNow.AddHours(-(_config.TempRetentionHours + 1));
        CreateTempFile("sess_expired12", expiredTime);

        var result = _manager.CleanupExpiredFiles();

        Assert.Equal(1, result.ScannedCount);
        Assert.Equal(1, result.DeletedCount);
    }

    [Fact]
    public void CleanupExpiredFiles_OrphanedFiles_ShouldCleanup()
    {
        // Create an orphaned file (document without metadata)
        var orphanedPath = Path.Combine(_tempDir, "aspose_session_orphan_20240101120000.docx");
        File.WriteAllText(orphanedPath, "orphaned content");
        File.SetLastWriteTimeUtc(orphanedPath, DateTime.UtcNow.AddHours(-(_config.TempRetentionHours + 1)));

        var result = _manager.CleanupExpiredFiles();

        Assert.Equal(1, result.DeletedCount);
        Assert.False(File.Exists(orphanedPath));
    }

    #endregion

    #region ListRecoverableFiles Tests

    [Fact]
    public void ListRecoverableFiles_EmptyDirectory_ShouldReturnEmpty()
    {
        var files = _manager.ListRecoverableFiles().ToList();

        Assert.Empty(files);
    }

    [Fact]
    public void ListRecoverableFiles_WithValidFiles_ShouldReturnFileInfo()
    {
        var sessionId = "sess_list1234";
        CreateTempFile(sessionId, DateTime.UtcNow);

        var files = _manager.ListRecoverableFiles().ToList();

        Assert.Single(files);
        Assert.Equal(sessionId, files[0].SessionId);
        Assert.Equal("Word", files[0].DocumentType);
    }

    [Fact]
    public void ListRecoverableFiles_ShouldOrderByMostRecent()
    {
        CreateTempFile("sess_older123", DateTime.UtcNow.AddHours(-2));
        CreateTempFile("sess_newer123", DateTime.UtcNow.AddHours(-1));

        var files = _manager.ListRecoverableFiles().ToList();

        Assert.Equal(2, files.Count);
        Assert.Equal("sess_newer123", files[0].SessionId);
        Assert.Equal("sess_older123", files[1].SessionId);
    }

    [Fact]
    public void ListRecoverableFiles_WithMissingDocument_ShouldNotInclude()
    {
        // Create only metadata file without document
        var metadataPath = Path.Combine(_tempDir, "aspose_session_sess_nodoc123_20240101120000.docx.meta.json");
        var metadata = new
        {
            SessionId = "sess_nodoc123",
            OriginalPath = "/test/doc.docx",
            TempPath = Path.Combine(_tempDir, "aspose_session_sess_nodoc123_20240101120000.docx"),
            DocumentType = "Word",
            SavedAt = DateTime.UtcNow,
            PromptOnReconnect = false
        };
        File.WriteAllText(metadataPath, JsonSerializer.Serialize(metadata));

        var files = _manager.ListRecoverableFiles().ToList();

        Assert.Empty(files);
    }

    #endregion

    #region RecoverSession Tests

    [Fact]
    public void RecoverSession_NonExistentSession_ShouldReturnError()
    {
        var result = _manager.RecoverSession("sess_nonexist");

        Assert.False(result.Success);
        Assert.Contains("No recoverable session found", result.ErrorMessage);
    }

    [Fact]
    public void RecoverSession_ValidSession_ShouldRecover()
    {
        var sessionId = "sess_recover1";
        CreateTempFile(sessionId, DateTime.UtcNow, "/original/path/doc.docx");

        var targetPath = Path.Combine(_tempDir, "recovered.docx");
        var result = _manager.RecoverSession(sessionId, targetPath);

        Assert.True(result.Success);
        Assert.Equal(targetPath, result.RecoveredPath);
        Assert.True(File.Exists(targetPath));
    }

    [Fact]
    public void RecoverSession_WithDeleteAfterRecover_ShouldDeleteTempFiles()
    {
        var sessionId = "sess_delafter";
        var (docPath, metaPath) = CreateTempFile(sessionId, DateTime.UtcNow);

        var targetPath = Path.Combine(_tempDir, "recovered_del.docx");
        var result = _manager.RecoverSession(sessionId, targetPath);

        Assert.True(result.Success);
        Assert.False(File.Exists(docPath));
        Assert.False(File.Exists(metaPath));
    }

    [Fact]
    public void RecoverSession_WithoutDeleteAfterRecover_ShouldKeepTempFiles()
    {
        var sessionId = "sess_keepafter";
        var (docPath, metaPath) = CreateTempFile(sessionId, DateTime.UtcNow);

        var targetPath = Path.Combine(_tempDir, "recovered_keep.docx");
        var result = _manager.RecoverSession(sessionId, targetPath, false);

        Assert.True(result.Success);
        Assert.True(File.Exists(docPath));
        Assert.True(File.Exists(metaPath));
    }

    #endregion

    #region DeleteTempSession Tests

    [Fact]
    public void DeleteTempSession_NonExistentSession_ShouldReturnFalse()
    {
        var result = _manager.DeleteTempSession("sess_nonexist");

        Assert.False(result);
    }

    [Fact]
    public void DeleteTempSession_ValidSession_ShouldDelete()
    {
        var sessionId = "sess_todel12";
        var (docPath, metaPath) = CreateTempFile(sessionId, DateTime.UtcNow);

        var result = _manager.DeleteTempSession(sessionId);

        Assert.True(result);
        Assert.False(File.Exists(docPath));
        Assert.False(File.Exists(metaPath));
    }

    #endregion

    #region GetStats Tests

    [Fact]
    public void GetStats_EmptyDirectory_ShouldReturnZeros()
    {
        var stats = _manager.GetStats();

        Assert.Equal(0, stats.TotalCount);
        Assert.Equal(0, stats.TotalSizeBytes);
        Assert.Equal(0, stats.ExpiredCount);
    }

    [Fact]
    public void GetStats_WithFiles_ShouldReturnCorrectCounts()
    {
        CreateTempFile("sess_stats123", DateTime.UtcNow);
        CreateTempFile("sess_stats456", DateTime.UtcNow.AddHours(-(_config.TempRetentionHours + 1)));

        var stats = _manager.GetStats();

        Assert.Equal(2, stats.TotalCount);
        Assert.True(stats.TotalSizeBytes > 0);
        Assert.Equal(1, stats.ExpiredCount);
    }

    #endregion

    #region IHostedService Tests

    [Fact]
    public async Task StartAsync_ShouldComplete()
    {
        var manager = new TempFileManager(_config);

        var exception = await Record.ExceptionAsync(async () =>
        {
            await manager.StartAsync(CancellationToken.None);
            await manager.StopAsync(CancellationToken.None);
        });

        Assert.Null(exception);
        manager.Dispose();
    }

    [Fact]
    public async Task StartAsync_WhenDisabled_ShouldSkip()
    {
        var disabledConfig = new SessionConfig
        {
            Enabled = false,
            TempDirectory = _tempDir
        };
        var manager = new TempFileManager(disabledConfig);

        var exception = await Record.ExceptionAsync(async () =>
        {
            await manager.StartAsync(CancellationToken.None);
            await manager.StopAsync(CancellationToken.None);
        });

        Assert.Null(exception);
        manager.Dispose();
    }

    #endregion
}
