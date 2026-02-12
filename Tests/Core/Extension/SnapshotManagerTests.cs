using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Extension.Transport;
using Microsoft.Extensions.Logging;
using Moq;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for SnapshotManager class.
/// </summary>
public class SnapshotManagerTests : IDisposable
{
    private readonly Mock<ILogger<SnapshotManager>> _loggerMock = new();
    private readonly SnapshotManager _manager;

    public SnapshotManagerTests()
    {
        var config = new ExtensionConfig { Enabled = true };
        _manager = new SnapshotManager(config, _loggerMock.Object);
    }

    public void Dispose()
    {
        _manager.Dispose();
        GC.SuppressFinalize(this);
    }

    #region Constructor Tests

    [Fact]
    public void Constructor_InitializesWithZeroPendingSnapshots()
    {
        Assert.Equal(0, _manager.PendingSnapshotCount);
    }

    #endregion

    #region Helper Methods

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

    #region StartAsync Tests

    [Fact]
    public async Task StartAsync_WhenDisabled_DoesNotStartCleanupTask()
    {
        var config = new ExtensionConfig { Enabled = false };
        using var manager = new SnapshotManager(config, _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        Assert.Equal(0, manager.PendingSnapshotCount);
    }

    [Fact]
    public async Task StartAsync_WhenEnabled_StartsSuccessfully()
    {
        await _manager.StartAsync(CancellationToken.None);

        Assert.Equal(0, _manager.PendingSnapshotCount);
    }

    #endregion

    #region StopAsync Tests

    [Fact]
    public async Task StopAsync_WithoutStart_DoesNotThrow()
    {
        var config = new ExtensionConfig { Enabled = false };
        using var manager = new SnapshotManager(config, _loggerMock.Object);

        var exception = await Record.ExceptionAsync(() => manager.StopAsync(CancellationToken.None));

        Assert.Null(exception);
    }

    [Fact]
    public async Task StopAsync_AfterStart_StopsSuccessfully()
    {
        await _manager.StartAsync(CancellationToken.None);

        var exception = await Record.ExceptionAsync(() => _manager.StopAsync(CancellationToken.None));

        Assert.Null(exception);
    }

    [Fact]
    public async Task StopAsync_CleansUpPendingSnapshots()
    {
        var transportMock = new Mock<IExtensionTransport>();
        _manager.RegisterTransport("ext1", transportMock.Object);

        await _manager.StartAsync(CancellationToken.None);

        var metadata = CreateTestMetadata();
        _manager.RecordSnapshot("ext1", metadata);

        Assert.Equal(1, _manager.PendingSnapshotCount);

        await _manager.StopAsync(CancellationToken.None);

        Assert.Equal(0, _manager.PendingSnapshotCount);
    }

    #endregion

    #region RegisterTransport Tests

    [Fact]
    public void RegisterTransport_AddsTransportSuccessfully()
    {
        var transportMock = new Mock<IExtensionTransport>();

        var exception = Record.Exception(() => _manager.RegisterTransport("ext1", transportMock.Object));

        Assert.Null(exception);
    }

    [Fact]
    public void RegisterTransport_SameExtensionTwice_OverwritesTransport()
    {
        var transport1 = new Mock<IExtensionTransport>();
        var transport2 = new Mock<IExtensionTransport>();

        _manager.RegisterTransport("ext1", transport1.Object);
        var exception = Record.Exception(() => _manager.RegisterTransport("ext1", transport2.Object));

        Assert.Null(exception);
    }

    #endregion

    #region UnregisterTransport Tests

    [Fact]
    public void UnregisterTransport_ExistingTransport_RemovesSuccessfully()
    {
        var transportMock = new Mock<IExtensionTransport>();
        _manager.RegisterTransport("ext1", transportMock.Object);

        var exception = Record.Exception(() => _manager.UnregisterTransport("ext1"));

        Assert.Null(exception);
    }

    [Fact]
    public void UnregisterTransport_NonExistentTransport_DoesNotThrow()
    {
        var exception = Record.Exception(() => _manager.UnregisterTransport("nonexistent"));

        Assert.Null(exception);
    }

    #endregion

    #region RecordSnapshot Tests

    [Fact]
    public void RecordSnapshot_IncreasesPendingCount()
    {
        var metadata = CreateTestMetadata();

        _manager.RecordSnapshot("ext1", metadata);

        Assert.Equal(1, _manager.PendingSnapshotCount);
    }

    [Fact]
    public void RecordSnapshot_MultipleDifferentExtensions_TracksEachSeparately()
    {
        var metadata1 = CreateTestMetadata();
        var metadata2 = CreateTestMetadata(2);

        _manager.RecordSnapshot("ext1", metadata1);
        _manager.RecordSnapshot("ext2", metadata2);

        Assert.Equal(2, _manager.PendingSnapshotCount);
        Assert.Equal(1, _manager.GetPendingSnapshotCount("ext1"));
        Assert.Equal(1, _manager.GetPendingSnapshotCount("ext2"));
    }

    [Fact]
    public void RecordSnapshot_SameExtensionMultipleSequences_TracksAll()
    {
        var metadata1 = CreateTestMetadata();
        var metadata2 = CreateTestMetadata(2);
        var metadata3 = CreateTestMetadata(3);

        _manager.RecordSnapshot("ext1", metadata1);
        _manager.RecordSnapshot("ext1", metadata2);
        _manager.RecordSnapshot("ext1", metadata3);

        Assert.Equal(3, _manager.PendingSnapshotCount);
        Assert.Equal(3, _manager.GetPendingSnapshotCount("ext1"));
    }

    #endregion

    #region HandleAck Tests

    [Fact]
    public void HandleAck_ExistingSnapshot_ReturnsTrueAndRemoves()
    {
        var metadata = CreateTestMetadata(42);
        _manager.RecordSnapshot("ext1", metadata);

        var result = _manager.HandleAck("ext1", 42);

        Assert.True(result);
        Assert.Equal(0, _manager.PendingSnapshotCount);
    }

    [Fact]
    public void HandleAck_NonExistentSnapshot_ReturnsFalse()
    {
        var result = _manager.HandleAck("ext1", 999);

        Assert.False(result);
    }

    [Fact]
    public void HandleAck_WrongExtension_ReturnsFalse()
    {
        var metadata = CreateTestMetadata(42);
        _manager.RecordSnapshot("ext1", metadata);

        var result = _manager.HandleAck("ext2", 42);

        Assert.False(result);
        Assert.Equal(1, _manager.PendingSnapshotCount);
    }

    [Fact]
    public void HandleAck_WithRegisteredTransport_CallsCleanup()
    {
        var transportMock = new Mock<IExtensionTransport>();
        _manager.RegisterTransport("ext1", transportMock.Object);

        var metadata = CreateTestMetadata(42);
        _manager.RecordSnapshot("ext1", metadata);

        _manager.HandleAck("ext1", 42);

        transportMock.Verify(t => t.Cleanup(metadata), Times.Once);
    }

    [Fact]
    public void HandleAck_WithoutRegisteredTransport_DoesNotThrow()
    {
        var metadata = CreateTestMetadata(42);
        _manager.RecordSnapshot("ext1", metadata);

        var exception = Record.Exception(() => _manager.HandleAck("ext1", 42));

        Assert.Null(exception);
    }

    [Fact]
    public void HandleAck_SameSequenceTwice_SecondReturnsFalse()
    {
        var metadata = CreateTestMetadata(42);
        _manager.RecordSnapshot("ext1", metadata);

        var result1 = _manager.HandleAck("ext1", 42);
        var result2 = _manager.HandleAck("ext1", 42);

        Assert.True(result1);
        Assert.False(result2);
    }

    #endregion

    #region GetPendingSnapshotCount Tests

    [Fact]
    public void GetPendingSnapshotCount_NoSnapshots_ReturnsZero()
    {
        var count = _manager.GetPendingSnapshotCount("ext1");

        Assert.Equal(0, count);
    }

    [Fact]
    public void GetPendingSnapshotCount_AfterAck_DecreasesCount()
    {
        var metadata1 = CreateTestMetadata();
        var metadata2 = CreateTestMetadata(2);

        _manager.RecordSnapshot("ext1", metadata1);
        _manager.RecordSnapshot("ext1", metadata2);

        Assert.Equal(2, _manager.GetPendingSnapshotCount("ext1"));

        _manager.HandleAck("ext1", 1);

        Assert.Equal(1, _manager.GetPendingSnapshotCount("ext1"));
    }

    #endregion

    #region CleanupExtensionSnapshots Tests

    [Fact]
    public void CleanupExtensionSnapshots_RemovesAllForExtension()
    {
        var metadata1 = CreateTestMetadata();
        var metadata2 = CreateTestMetadata(2);
        var metadata3 = CreateTestMetadata(3);

        _manager.RecordSnapshot("ext1", metadata1);
        _manager.RecordSnapshot("ext1", metadata2);
        _manager.RecordSnapshot("ext2", metadata3);

        _manager.CleanupExtensionSnapshots("ext1");

        Assert.Equal(1, _manager.PendingSnapshotCount);
        Assert.Equal(0, _manager.GetPendingSnapshotCount("ext1"));
        Assert.Equal(1, _manager.GetPendingSnapshotCount("ext2"));
    }

    [Fact]
    public void CleanupExtensionSnapshots_CallsTransportCleanup()
    {
        var transportMock = new Mock<IExtensionTransport>();
        _manager.RegisterTransport("ext1", transportMock.Object);

        var metadata1 = CreateTestMetadata();
        var metadata2 = CreateTestMetadata(2);

        _manager.RecordSnapshot("ext1", metadata1);
        _manager.RecordSnapshot("ext1", metadata2);

        _manager.CleanupExtensionSnapshots("ext1");

        transportMock.Verify(t => t.Cleanup(It.IsAny<ExtensionMetadata>()), Times.Exactly(2));
    }

    [Fact]
    public void CleanupExtensionSnapshots_NonExistentExtension_DoesNotThrow()
    {
        var exception = Record.Exception(() => _manager.CleanupExtensionSnapshots("nonexistent"));

        Assert.Null(exception);
    }

    #endregion

    #region Dispose Tests

    [Fact]
    public void Dispose_CleansUpAllSnapshots()
    {
        var transportMock = new Mock<IExtensionTransport>();
        _manager.RegisterTransport("ext1", transportMock.Object);

        var metadata = CreateTestMetadata();
        _manager.RecordSnapshot("ext1", metadata);

        _manager.Dispose();

        transportMock.Verify(t => t.Cleanup(metadata), Times.Once);
    }

    [Fact]
    public void Dispose_MultipleTimes_DoesNotThrow()
    {
        _manager.Dispose();

        var exception = Record.Exception(() => _manager.Dispose());

        Assert.Null(exception);
    }

    #endregion
}
