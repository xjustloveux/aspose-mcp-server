using System.Diagnostics;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Extension.Transport;

namespace AsposeMcpServer.Tests.Core.Extension.Transport;

/// <summary>
///     Unit tests for StdinTransport class.
/// </summary>
public class StdinTransportTests
{
    private readonly StdinTransport _transport = new();

    #region Mode Tests

    [Fact]
    public void Mode_ReturnsCorrectValue()
    {
        Assert.Equal("stdin", _transport.Mode);
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
    public async Task SendAsync_Success_SetsMetadataProperties()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4, 5 };

        var result = await _transport.SendAsync(process, data, metadata);

        Assert.True(result);
        Assert.Equal("stdin", metadata.TransportMode);
        Assert.Equal(5, metadata.DataSize);
    }

    [Fact]
    public async Task SendAsync_Success_DoesNotModifyMmapNameOrFilePath()
    {
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);

        Assert.Null(metadata.MmapName);
        Assert.Null(metadata.FilePath);
    }

    #endregion

    #region Cleanup Tests

    [Fact]
    public void Cleanup_DoesNotThrow()
    {
        var metadata = CreateTestMetadata();

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [Fact]
    public void Cleanup_WithNullMetadata_DoesNotThrow()
    {
        var metadata = new ExtensionMetadata();

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [Fact]
    public void Cleanup_CanBeCalledMultipleTimes()
    {
        var metadata = CreateTestMetadata();

        var exception = Record.Exception(() =>
        {
            _transport.Cleanup(metadata);
            _transport.Cleanup(metadata);
            _transport.Cleanup(metadata);
        });

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
