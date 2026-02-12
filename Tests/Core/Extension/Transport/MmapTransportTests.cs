using System.Diagnostics;
using System.IO.MemoryMappedFiles;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Extension.Transport;

// MemoryMappedFile.OpenExisting is Windows-only, but tests are conditionally skipped on non-Windows platforms
#pragma warning disable CA1416

namespace AsposeMcpServer.Tests.Core.Extension.Transport;

/// <summary>
///     Unit tests for MmapTransport class.
///     Tests are cross-platform compatible.
/// </summary>
public class MmapTransportTests : IDisposable
{
    private readonly string _tempDirectory;
    private readonly MmapTransport _transport;

    public MmapTransportTests()
    {
        _tempDirectory = Path.Combine(Path.GetTempPath(), $"MmapTransportTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDirectory);
        _transport = new MmapTransport(tempDirectory: _tempDirectory);
    }

    public void Dispose()
    {
        _transport.Dispose();
        try
        {
            if (Directory.Exists(_tempDirectory))
                Directory.Delete(_tempDirectory, true);
        }
// ReSharper disable once EmptyGeneralCatchClause
        catch
        {
        }

        GC.SuppressFinalize(this);
    }

    #region Mode Tests

    [Fact]
    public void Mode_ReturnsCorrectValue()
    {
        Assert.Equal("mmap", _transport.Mode);
    }

    #endregion

    #region SendAsync Tests

    [Fact]
    public async Task SendAsync_ProcessHasExited_ReturnsFalse()
    {
        using var process = new Process();
        process.StartInfo = CreateExitingProcessStartInfo();
        process.Start();
        await process.WaitForExitAsync();

        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        var result = await _transport.SendAsync(process, data, metadata);

        Assert.False(result);
    }

    [Fact]
    public async Task SendAsync_DisposedTransport_ThrowsObjectDisposedException()
    {
        var transport = new MmapTransport(tempDirectory: _tempDirectory);
        transport.Dispose();

        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await Assert.ThrowsAsync<ObjectDisposedException>(() =>
            transport.SendAsync(process, data, metadata));
    }

    [Fact]
    public async Task SendAsync_Success_SetsMetadataProperties()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4, 5 };

        var result = await _transport.SendAsync(process, data, metadata);

        Assert.True(result);
        Assert.Equal("mmap", metadata.TransportMode);
        Assert.Equal(5, metadata.DataSize);
        Assert.NotNull(metadata.MmapName);
        Assert.Contains("aspose_", metadata.MmapName);

        _transport.Cleanup(metadata);
    }

    [Fact]
    public async Task SendAsync_Success_CreatesMemoryMappedFile()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var testData = new byte[] { 10, 20, 30, 40, 50 };

        var result = await _transport.SendAsync(process, testData, metadata);

        Assert.True(result);
        Assert.NotNull(metadata.MmapName);

        if (OperatingSystem.IsMacOS())
        {
            Assert.NotNull(metadata.FilePath);
            Assert.True(File.Exists(metadata.FilePath));
            var fileData = await File.ReadAllBytesAsync(metadata.FilePath);
            Assert.Equal(testData, fileData);
        }
        else
        {
            using var mmf = MemoryMappedFile.OpenExisting(metadata.MmapName);
            using var accessor = mmf.CreateViewAccessor(0, testData.Length, MemoryMappedFileAccess.Read);
            var readData = new byte[testData.Length];
            accessor.ReadArray(0, readData, 0, testData.Length);
            Assert.Equal(testData, readData);
        }

        _transport.Cleanup(metadata);
    }

    [Fact]
    public async Task SendAsync_MacOS_SetsFilePath()
    {
        if (!OperatingSystem.IsMacOS())
            return;

        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var testData = new byte[] { 1, 2, 3 };

        var result = await _transport.SendAsync(process, testData, metadata);

        Assert.True(result);
        Assert.NotNull(metadata.FilePath);
        Assert.True(File.Exists(metadata.FilePath));

        _transport.Cleanup(metadata);

        await Task.Delay(1500);
        Assert.False(File.Exists(metadata.FilePath));
    }

    #endregion

    #region Cleanup Tests

    [Fact]
    public void Cleanup_NullMmapName_DoesNotThrow()
    {
        var metadata = CreateTestMetadata();
        metadata.MmapName = null;

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [Fact]
    public void Cleanup_EmptyMmapName_DoesNotThrow()
    {
        var metadata = CreateTestMetadata();
        metadata.MmapName = string.Empty;

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [Fact]
    public async Task Cleanup_ExistingMmap_SchedulesDelayedDispose()
    {
        if (OperatingSystem.IsMacOS())
            return;

        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);
        var mmapName = metadata.MmapName;

        _transport.Cleanup(metadata);

        await Task.Delay(2500);

        Assert.Throws<FileNotFoundException>(() =>
            MemoryMappedFile.OpenExisting(mmapName!));
    }

    [Fact]
    public void Cleanup_AlreadyCleanedUp_DoesNotThrow()
    {
        var metadata = CreateTestMetadata();
        metadata.MmapName = "nonexistent_mmap";

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    #endregion

    #region Dispose Tests

    [Fact]
    public async Task Dispose_CleansUpAllMmaps()
    {
        if (OperatingSystem.IsMacOS())
            return;

        var transport = new MmapTransport(tempDirectory: _tempDirectory);
        using var process = CreateTestProcess();
        var metadata1 = CreateTestMetadata();
        var metadata2 = CreateTestMetadata(2);
        var data = new byte[] { 1, 2, 3, 4 };

        await transport.SendAsync(process, data, metadata1);
        await transport.SendAsync(process, data, metadata2);

        transport.Dispose();

        Assert.Throws<FileNotFoundException>(() =>
            MemoryMappedFile.OpenExisting(metadata1.MmapName!));
        Assert.Throws<FileNotFoundException>(() =>
            MemoryMappedFile.OpenExisting(metadata2.MmapName!));
    }

    [Fact]
    public async Task Dispose_MacOS_CleansUpAllFiles()
    {
        if (!OperatingSystem.IsMacOS())
            return;

        var transport = new MmapTransport(tempDirectory: _tempDirectory);
        using var process = CreateTestProcess();
        var metadata1 = CreateTestMetadata();
        var metadata2 = CreateTestMetadata(2);
        var data = new byte[] { 1, 2, 3, 4 };

        await transport.SendAsync(process, data, metadata1);
        await transport.SendAsync(process, data, metadata2);

        var filePath1 = metadata1.FilePath;
        var filePath2 = metadata2.FilePath;

        Assert.True(File.Exists(filePath1));
        Assert.True(File.Exists(filePath2));

        transport.Dispose();

        Assert.False(File.Exists(filePath1));
        Assert.False(File.Exists(filePath2));
    }

    [Fact]
    public void Dispose_CanBeCalledMultipleTimes()
    {
        var transport = new MmapTransport(tempDirectory: _tempDirectory);

        var exception = Record.Exception(() =>
        {
            transport.Dispose();
            transport.Dispose();
        });

        Assert.Null(exception);
    }

    #endregion

    #region ForceCleanup Tests

    [Fact]
    public async Task ForceCleanup_ExistingMmap_ImmediatelyDisposes()
    {
        if (OperatingSystem.IsMacOS())
            return;

        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);
        var mmapName = metadata.MmapName!;

        var result = _transport.ForceCleanup(mmapName);

        Assert.True(result);
        Assert.Throws<FileNotFoundException>(() =>
            MemoryMappedFile.OpenExisting(mmapName));
    }

    [Fact]
    public async Task ForceCleanup_MacOS_ImmediatelyDeletesFile()
    {
        if (!OperatingSystem.IsMacOS())
            return;

        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);
        var mmapName = metadata.MmapName!;
        var filePath = metadata.FilePath!;

        Assert.True(File.Exists(filePath));

        var result = _transport.ForceCleanup(mmapName);

        Assert.True(result);
        Assert.False(File.Exists(filePath));
    }

    [Fact]
    public void ForceCleanup_NonexistentMmap_ReturnsFalse()
    {
        var result = _transport.ForceCleanup("nonexistent_mmap_name");

        Assert.False(result);
    }

    [Fact]
    public void ForceCleanup_NullMmapName_ReturnsFalse()
    {
        var result = _transport.ForceCleanup(null!);

        Assert.False(result);
    }

    [Fact]
    public void ForceCleanup_EmptyMmapName_ReturnsFalse()
    {
        var result = _transport.ForceCleanup(string.Empty);

        Assert.False(result);
    }

    #endregion

    #region Diagnostic Properties Tests

    [Fact]
    public void ActiveMmapCount_InitiallyZero()
    {
        Assert.Equal(0, _transport.ActiveMmapCount);
    }

    [Fact]
    public async Task ActiveMmapCount_IncreasesAfterSend()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);

        Assert.Equal(1, _transport.ActiveMmapCount);

        _transport.Cleanup(metadata);
    }

    [Fact]
    public async Task ActiveMmapCount_DecreasesAfterCleanup()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);
        Assert.Equal(1, _transport.ActiveMmapCount);

        _transport.Cleanup(metadata);
        Assert.Equal(0, _transport.ActiveMmapCount);
    }

    [Fact]
    public void PendingCleanupCount_InitiallyZero()
    {
        Assert.Equal(0, _transport.PendingCleanupCount);
    }

    [Fact]
    public async Task PendingCleanupCount_IncreasesAfterCleanup()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);
        _transport.Cleanup(metadata);

        Assert.True(_transport.PendingCleanupCount >= 1);
    }

    #endregion

    #region Helper Methods

    private static ProcessStartInfo CreateExitingProcessStartInfo()
    {
        return OperatingSystem.IsWindows()
            ? new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = "/c echo test",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            }
            : new ProcessStartInfo
            {
                FileName = "/bin/sh",
                Arguments = "-c \"echo test\"",
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };
    }

    private static Process CreateTestProcess()
    {
        var process = new Process
        {
            StartInfo = OperatingSystem.IsWindows()
                ? new ProcessStartInfo
                {
                    FileName = "findstr",
                    Arguments = "x",
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
                : new ProcessStartInfo
                {
                    FileName = "cat",
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
        };
        process.Start();
        return process;
    }

    private static ExtensionMetadata CreateTestMetadata(long sequenceNumber = 1)
    {
        return new ExtensionMetadata
        {
            SessionId = $"test_session_{Guid.NewGuid():N}",
            SequenceNumber = sequenceNumber,
            DocumentType = "word",
            OutputFormat = "pdf",
            MimeType = "application/pdf"
        };
    }

    #endregion
}
