using System.Diagnostics;
using System.Runtime.Versioning;
using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Extension.Transport;
using Microsoft.Extensions.Logging;
using Moq;

namespace AsposeMcpServer.Tests.Core.Extension.Transport;

/// <summary>
///     Unit tests for StdinTransport class.
/// </summary>
[SupportedOSPlatform("windows")]
public class StdinTransportTests
{
    private readonly Mock<ILogger<StdinTransport>> _loggerMock = new();
    private readonly StdinTransport _transport = new();

    #region Mode Tests

    [SkippableFact]
    public void Mode_ReturnsCorrectValue()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        Assert.Equal("stdin", _transport.Mode);
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
    public async Task SendAsync_Success_SetsMetadataProperties()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4, 5 };

        var result = await _transport.SendAsync(process, data, metadata);

        Assert.True(result);
        Assert.Equal("stdin", metadata.TransportMode);
        Assert.Equal(5, metadata.DataSize);
    }

    [SkippableFact]
    public async Task SendAsync_Success_DoesNotModifyMmapNameOrFilePath()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[] { 1, 2, 3, 4 };

        await _transport.SendAsync(process, data, metadata);

        Assert.Null(metadata.MmapName);
        Assert.Null(metadata.FilePath);
    }

    #endregion

    #region Cleanup Tests

    [SkippableFact]
    public void Cleanup_DoesNotThrow()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var metadata = CreateTestMetadata();

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [SkippableFact]
    public void Cleanup_WithNullMetadata_DoesNotThrow()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var metadata = new ExtensionMetadata();

        var exception = Record.Exception(() => _transport.Cleanup(metadata));

        Assert.Null(exception);
    }

    [SkippableFact]
    public void Cleanup_CanBeCalledMultipleTimes()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
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

    #region Constructor Tests

    [SkippableFact]
    public void Constructor_WithDefaultParameters_DoesNotThrow()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var exception = Record.Exception(() => new StdinTransport());

        Assert.Null(exception);
    }

    [SkippableFact]
    public void Constructor_WithLogger_DoesNotThrow()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var exception = Record.Exception(() => new StdinTransport(_loggerMock.Object));

        Assert.Null(exception);
    }

    [SkippableFact]
    public void Constructor_WithCustomParameters_DoesNotThrow()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var exception = Record.Exception(() => new StdinTransport(
            _loggerMock.Object,
            60000,
            50 * 1024 * 1024));

        Assert.Null(exception);
    }

    #endregion

    #region DataSize Limit Tests

    [SkippableFact]
    public async Task SendAsync_DataExceedsMaxSize_ReturnsFalse()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var transport = new StdinTransport(_loggerMock.Object, maxDataSize: 10);
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var largeData = new byte[20];

        var result = await transport.SendAsync(process, largeData, metadata);

        Assert.False(result);
    }

    [SkippableFact]
    public async Task SendAsync_DataAtMaxSize_ReturnsTrue()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var transport = new StdinTransport(_loggerMock.Object, maxDataSize: 100);
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[100];

        var result = await transport.SendAsync(process, data, metadata);

        Assert.True(result);
    }

    [SkippableFact]
    public async Task SendAsync_DataBelowMaxSize_ReturnsTrue()
    {
        Skip.IfNot(OperatingSystem.IsWindows(), "Only supported on Windows");
        var transport = new StdinTransport(_loggerMock.Object, maxDataSize: 100);
        using var process = CreateTestProcess();
        var metadata = CreateTestMetadata();
        var data = new byte[50];

        var result = await transport.SendAsync(process, data, metadata);

        Assert.True(result);
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
