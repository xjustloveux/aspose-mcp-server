using System.Diagnostics;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Extension.Transport;

namespace AsposeMcpServer.Tests.Core.Extension.Transport;

/// <summary>
///     Unit tests for FileTransport class.
/// </summary>
public class FileTransportTests : IDisposable
{
    private readonly string _tempDirectory;
    private readonly FileTransport _transport;

    public FileTransportTests()
    {
        _tempDirectory = Path.Combine(Path.GetTempPath(), $"FileTransportTests_{Guid.NewGuid():N}");
        _transport = new FileTransport(_tempDirectory);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
                Directory.Delete(_tempDirectory, true);
        }
        // ReSharper disable once EmptyGeneralCatchClause - Best-effort cleanup in dispose
        catch
        {
        }

        GC.SuppressFinalize(this);
    }

    #region Mode Tests

    [Fact]
    public void Mode_ReturnsCorrectValue()
    {
        Assert.Equal("file", _transport.Mode);
    }

    #endregion

    #region Constructor Tests

    [Fact]
    public void Constructor_CreatesTempDirectory()
    {
        Assert.True(Directory.Exists(_tempDirectory));
    }

    [Fact]
    public void Constructor_ExistingDirectory_DoesNotThrow()
    {
        var existingDir = Path.Combine(Path.GetTempPath(), $"FileTransportTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(existingDir);

        try
        {
            var exception = Record.Exception(() => new FileTransport(existingDir));
            Assert.Null(exception);
        }
        finally
        {
            Directory.Delete(existingDir, true);
        }
    }

    [Fact]
    public void Constructor_NullTempDirectory_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => new FileTransport(null!));
    }

    [Fact]
    public void Constructor_EmptyTempDirectory_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => new FileTransport(string.Empty));
    }

    [Fact]
    public void Constructor_WhitespaceTempDirectory_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => new FileTransport("   "));
    }

    #endregion

    #region SendAsync Tests

    [Fact]
    public async Task SendAsync_ProcessHasExited_ReturnsFalse()
    {
        using var process = new Process();
        process.StartInfo = new ProcessStartInfo
        {
            FileName = "cmd",
            Arguments = "/c echo test",
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        };
        process.Start();
        await process.WaitForExitAsync();

        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        var result = await _transport.SendAsync(process, data, metadata);

        Assert.False(result);
    }

    [Fact]
    public async Task SendAsync_Success_CreatesFile()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var testData = new byte[] { 10, 20, 30, 40, 50 };

        var result = await _transport.SendAsync(process, testData, metadata);

        Assert.True(result);
        Assert.NotNull(metadata.FilePath);
        Assert.True(File.Exists(metadata.FilePath));

        var readData = await File.ReadAllBytesAsync(metadata.FilePath);
        Assert.Equal(testData, readData);

        _transport.Cleanup(metadata);
    }

    [Fact]
    public async Task SendAsync_Success_SetsMetadataProperties()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4, 5 };

        var result = await _transport.SendAsync(process, data, metadata);

        Assert.True(result);
        Assert.Equal("file", metadata.TransportMode);
        Assert.Equal(5, metadata.DataSize);
        Assert.NotNull(metadata.FilePath);
        Assert.Contains(metadata.SessionId, metadata.FilePath);
        Assert.Contains(metadata.OutputFormat, metadata.FilePath);

        _transport.Cleanup(metadata);
    }

    [Fact]
    public async Task SendAsync_Success_FileNameIncludesSequenceNumber()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata(42);
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);

        Assert.Contains("42", metadata.FilePath);

        _transport.Cleanup(metadata);
    }

    #endregion

    #region Cleanup Tests

    [Fact]
    public void Cleanup_NullFilePath_DoesNotThrow()
    {
        var metadata = CreateTestMetadata();
        metadata.FilePath = null;

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [Fact]
    public void Cleanup_EmptyFilePath_DoesNotThrow()
    {
        var metadata = CreateTestMetadata();
        metadata.FilePath = string.Empty;

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [Fact]
    public async Task Cleanup_ExistingFile_DeletesFile()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);
        var filePath = metadata.FilePath;

        Assert.True(File.Exists(filePath));

        _transport.Cleanup(metadata);

        Assert.False(File.Exists(filePath));
    }

    [Fact]
    public void Cleanup_NonExistentFile_DoesNotThrow()
    {
        var metadata = CreateTestMetadata();
        metadata.FilePath = Path.Combine(_tempDirectory, "nonexistent.pdf");

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    #endregion

    #region Helper Methods

    private static Process CreateTestProcess()
    {
        var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "findstr",
                Arguments = "x",
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
