using System.Diagnostics;
using System.Runtime.Versioning;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Extension.Transport;

namespace AsposeMcpServer.Tests.Core.Extension.Transport;

/// <summary>
///     Unit tests for FileTransport class.
/// </summary>
[SupportedOSPlatform("windows")]
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
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
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

    [SkippableFact]
    public void Mode_ReturnsCorrectValue()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        Assert.Equal("file", _transport.Mode);
    }

    #endregion

    #region Constructor Tests

    [SkippableFact]
    public void Constructor_CreatesTempDirectory()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        Assert.True(Directory.Exists(_tempDirectory));
    }

    [SkippableFact]
    public void Constructor_ExistingDirectory_DoesNotThrow()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
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

    [SkippableFact]
    public void Constructor_NullTempDirectory_ThrowsArgumentException()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        Assert.Throws<ArgumentException>(() => new FileTransport(null!));
    }

    [SkippableFact]
    public void Constructor_EmptyTempDirectory_ThrowsArgumentException()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        Assert.Throws<ArgumentException>(() => new FileTransport(string.Empty));
    }

    [SkippableFact]
    public void Constructor_WhitespaceTempDirectory_ThrowsArgumentException()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        Assert.Throws<ArgumentException>(() => new FileTransport("   "));
    }

    #endregion

    #region SendAsync Tests

    [SkippableFact]
    public async Task SendAsync_ProcessHasExited_ReturnsFalse()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
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

    [SkippableFact]
    public async Task SendAsync_Success_CreatesFile()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
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

    [SkippableFact]
    public async Task SendAsync_Success_SetsMetadataProperties()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
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

    [SkippableFact]
    public async Task SendAsync_Success_FileNameIncludesSequenceNumber()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata(42);
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);

        Assert.Contains("42", metadata.FilePath);

        _transport.Cleanup(metadata);
    }

    #endregion

    #region Cleanup Tests

    [SkippableFact]
    public void Cleanup_NullFilePath_DoesNotThrow()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var metadata = CreateTestMetadata();
        metadata.FilePath = null;

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [SkippableFact]
    public void Cleanup_EmptyFilePath_DoesNotThrow()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var metadata = CreateTestMetadata();
        metadata.FilePath = string.Empty;

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [SkippableFact]
    public async Task Cleanup_ExistingFile_DeletesFile()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);
        var filePath = metadata.FilePath;

        Assert.True(File.Exists(filePath));

        _transport.Cleanup(metadata);

        Assert.False(File.Exists(filePath));
    }

    [SkippableFact]
    public void Cleanup_NonExistentFile_DoesNotThrow()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
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
