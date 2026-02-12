using AsposeMcpServer.Core.Extension;
using AsposeMcpServer.Core.Extension.Transport;
using Microsoft.Extensions.Logging;
using Moq;
using ExtensionClass = AsposeMcpServer.Core.Extension.Extension;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for the Extension class.
/// </summary>
public class ExtensionTests : IDisposable
{
    private readonly ExtensionConfig _config;
    private readonly ExtensionDefinition _definition;
    private readonly Mock<ILogger<ExtensionClass>> _loggerMock;
    private readonly SnapshotManager _snapshotManager;
    private readonly Mock<IExtensionTransport> _transportMock;

    public ExtensionTests()
    {
        _loggerMock = new Mock<ILogger<ExtensionClass>>();
        _transportMock = new Mock<IExtensionTransport>();
        var snapshotLoggerMock = new Mock<ILogger<SnapshotManager>>();
        _config = new ExtensionConfig
        {
            Enabled = true,
            MaxRestartAttempts = 3,
            RestartCooldownSeconds = 1,
            HealthCheckIntervalSeconds = 30
        };
        _snapshotManager = new SnapshotManager(_config, snapshotLoggerMock.Object);
        _definition = CreateTestDefinition();
    }

    public void Dispose()
    {
        _snapshotManager.Dispose();
        GC.SuppressFinalize(this);
    }

    #region StateChanged Event Tests

    [Fact]
    public async Task StateChanged_FiresOnStateTransition()
    {
        await using var extension = CreateExtensionWithInvalidCommand();
        var stateChanges = new List<ExtensionState>();
        extension.StateChanged += (_, state) => stateChanges.Add(state);

        await extension.EnsureStartedAsync();
        await Task.Delay(100);

        Assert.Contains(ExtensionState.Starting, stateChanges);
    }

    #endregion

    #region Constructor Tests

    [Fact]
    public async Task Constructor_InitializesWithUnloadedState()
    {
        await using var extension = CreateExtension();

        Assert.Equal(ExtensionState.Unloaded, extension.State);
    }

    [Fact]
    public async Task Constructor_InitializesDefinition()
    {
        await using var extension = CreateExtension();

        Assert.Equal(_definition.Id, extension.Definition.Id);
        Assert.Equal(_definition.DisplayName, extension.Definition.DisplayName);
    }

    [Fact]
    public async Task Constructor_InitializesRestartCountToZero()
    {
        await using var extension = CreateExtension();

        Assert.Equal(0, extension.RestartCount);
    }

    #endregion

    #region CanHandle Tests

    [Fact]
    public async Task CanHandle_SupportedDocumentTypeAndFormat_ReturnsTrue()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test",
            IsAvailable = true,
            SupportedDocumentTypes = ["word", "excel"],
            InputFormats = ["pdf", "html"]
        };
        await using var extension = CreateExtension(definition);

        Assert.True(extension.CanHandle("word", "pdf"));
        Assert.True(extension.CanHandle("excel", "html"));
    }

    [Fact]
    public async Task CanHandle_UnsupportedDocumentType_ReturnsFalse()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test",
            IsAvailable = true,
            SupportedDocumentTypes = ["word"],
            InputFormats = ["pdf"]
        };
        await using var extension = CreateExtension(definition);

        Assert.False(extension.CanHandle("powerpoint", "pdf"));
    }

    [Fact]
    public async Task CanHandle_UnsupportedFormat_ReturnsFalse()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test",
            IsAvailable = true,
            SupportedDocumentTypes = ["word"],
            InputFormats = ["pdf"]
        };
        await using var extension = CreateExtension(definition);

        Assert.False(extension.CanHandle("word", "png"));
    }

    [Fact]
    public async Task CanHandle_EmptyDocumentTypes_AcceptsAll()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test",
            IsAvailable = true,
            SupportedDocumentTypes = [],
            InputFormats = ["pdf"]
        };
        await using var extension = CreateExtension(definition);

        Assert.True(extension.CanHandle("word", "pdf"));
        Assert.True(extension.CanHandle("excel", "pdf"));
    }

    [Fact]
    public async Task CanHandle_EmptyInputFormats_AcceptsAll()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test",
            IsAvailable = true,
            SupportedDocumentTypes = ["word"],
            InputFormats = []
        };
        await using var extension = CreateExtension(definition);

        Assert.True(extension.CanHandle("word", "pdf"));
        Assert.True(extension.CanHandle("word", "html"));
    }

    [Fact]
    public async Task CanHandle_NotAvailable_ReturnsFalse()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test",
            IsAvailable = false,
            SupportedDocumentTypes = ["word"],
            InputFormats = ["pdf"]
        };
        await using var extension = CreateExtension(definition);

        Assert.False(extension.CanHandle("word", "pdf"));
    }

    [Fact]
    public async Task CanHandle_CaseInsensitive_ReturnsTrue()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test",
            IsAvailable = true,
            SupportedDocumentTypes = ["Word"],
            InputFormats = ["PDF"]
        };
        await using var extension = CreateExtension(definition);

        Assert.True(extension.CanHandle("WORD", "pdf"));
        Assert.True(extension.CanHandle("word", "PDF"));
    }

    #endregion

    #region EnsureStartedAsync Tests

    [Fact]
    public async Task EnsureStartedAsync_AfterDispose_ReturnsFalse()
    {
        var extension = CreateExtension();
        await extension.DisposeAsync();

        var result = await extension.EnsureStartedAsync();

        Assert.False(result);
    }

    [Fact]
    public async Task EnsureStartedAsync_InErrorState_ReturnsFalse()
    {
        await using var extension = CreateExtensionWithInvalidCommand();

        await extension.EnsureStartedAsync();
        await Task.Delay(100);

        var result = await extension.EnsureStartedAsync();

        Assert.False(result);
    }

    #endregion

    #region SendSnapshotAsync Tests

    [Fact]
    public async Task SendSnapshotAsync_AfterDispose_ReturnsFalse()
    {
        var extension = CreateExtension();
        await extension.DisposeAsync();

        var metadata = CreateTestMetadata();
        var result = await extension.SendSnapshotAsync([1, 2, 3], metadata);

        Assert.False(result);
    }

    [Fact]
    public async Task SendSnapshotAsync_WhenNotStarted_EnsuresStarted()
    {
        await using var extension = CreateExtensionWithInvalidCommand();
        var metadata = CreateTestMetadata();

        var result = await extension.SendSnapshotAsync([1, 2, 3], metadata);

        Assert.False(result);
    }

    #endregion

    #region SendHeartbeatAsync Tests

    [Fact]
    public async Task SendHeartbeatAsync_AfterDispose_ReturnsFalse()
    {
        var extension = CreateExtension();
        await extension.DisposeAsync();

        var result = await extension.SendHeartbeatAsync();

        Assert.False(result);
    }

    [Fact]
    public async Task SendHeartbeatAsync_NotIdleAndHeartbeatNotSupported_ReturnsFalse()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test",
            IsAvailable = true,
            Capabilities = new ExtensionCapabilities
            {
                SupportsHeartbeat = false
            }
        };
        await using var extension = CreateExtension(definition);

        var result = await extension.SendHeartbeatAsync();

        Assert.False(result);
    }

    [Fact]
    public async Task SendHeartbeatAsync_NotIdle_ReturnsFalse()
    {
        await using var extension = CreateExtension();

        var result = await extension.SendHeartbeatAsync();

        Assert.False(result);
    }

    #endregion

    #region SendCommandAsync Tests

    [Fact]
    public async Task SendCommandAsync_AfterDispose_ReturnsFailure()
    {
        var extension = CreateExtension();
        await extension.DisposeAsync();

        var result = await extension.SendCommandAsync("session1", "highlight");

        Assert.False(result.IsSuccess);
        Assert.Contains("not running", result.Error);
    }

    [Fact]
    public async Task SendCommandAsync_NotRunning_ReturnsFailure()
    {
        await using var extension = CreateExtension();

        var result = await extension.SendCommandAsync("session1", "highlight");

        Assert.False(result.IsSuccess);
    }

    #endregion

    #region HandleAck Tests

    [Fact]
    public async Task HandleAck_UpdatesLastActivity()
    {
        await using var extension = CreateExtension();
        var beforeAck = extension.LastActivity;

        Thread.Sleep(10);
        extension.HandleAck(1);

        Assert.True(extension.LastActivity > beforeAck);
    }

    [Fact]
    public async Task HandleAck_WithErrorStatus_DoesNotThrow()
    {
        await using var extension = CreateExtension();

        var exception = Record.Exception(() => extension.HandleAck(1, "failed", "Test error"));

        Assert.Null(exception);
    }

    [Fact]
    public async Task HandleAck_WithProcessedStatus_DoesNotThrow()
    {
        await using var extension = CreateExtension();

        var exception = Record.Exception(() => extension.HandleAck(1, "processed"));

        Assert.Null(exception);
    }

    #endregion

    #region StopAsync Tests

    [Fact]
    public async Task StopAsync_WhenNotStarted_CompletesWithoutError()
    {
        await using var extension = CreateExtension();

        var exception = await Record.ExceptionAsync(() => extension.StopAsync());

        Assert.Null(exception);
    }

    [Fact]
    public async Task StopAsync_WithResetRestartCount_WhenNotStarted_CompletesWithoutError()
    {
        await using var extension = CreateExtension();

        var exception = await Record.ExceptionAsync(() => extension.StopAsync(true));

        Assert.Null(exception);
    }

    [Fact]
    public async Task StopAsync_WithoutResetRestartCount_WhenNotStarted_CompletesWithoutError()
    {
        await using var extension = CreateExtension();

        var exception = await Record.ExceptionAsync(() => extension.StopAsync());

        Assert.Null(exception);
    }

    #endregion

    #region TryRestartAsync Tests

    [Fact]
    public async Task TryRestartAsync_AfterDispose_ReturnsFalse()
    {
        var extension = CreateExtension();
        await extension.DisposeAsync();

        var result = await extension.TryRestartAsync();

        Assert.False(result);
    }

    [Fact]
    public async Task TryRestartAsync_ConcurrentCalls_AtMostOneProceeds()
    {
        await using var extension = CreateExtensionWithInvalidCommand();

        var task1 = extension.TryRestartAsync();
        var task2 = extension.TryRestartAsync();

        var results = await Task.WhenAll(task1, task2);

        var successCount = results.Count(r => r);
        Assert.True(successCount <= 1, $"Expected at most 1 restart to proceed, but {successCount} did");
    }

    #endregion

    #region TryRecoverFromErrorAsync Tests

    [Fact]
    public async Task TryRecoverFromErrorAsync_AfterDispose_ReturnsFalse()
    {
        var extension = CreateExtension();
        await extension.DisposeAsync();

        var result = await extension.TryRecoverFromErrorAsync();

        Assert.False(result);
    }

    [Fact]
    public async Task TryRecoverFromErrorAsync_NotInErrorState_ReturnsFalse()
    {
        await using var extension = CreateExtension();

        var result = await extension.TryRecoverFromErrorAsync();

        Assert.False(result);
    }

    #endregion

    #region NotifySessionClosedAsync Tests

    [Fact]
    public async Task NotifySessionClosedAsync_AfterDispose_CompletesWithoutError()
    {
        var extension = CreateExtension();
        await extension.DisposeAsync();

        var exception = await Record.ExceptionAsync(() =>
            extension.NotifySessionClosedAsync("session1", null));

        Assert.Null(exception);
    }

    [Fact]
    public async Task NotifySessionClosedAsync_NotStarted_CompletesWithoutError()
    {
        await using var extension = CreateExtension();

        var exception = await Record.ExceptionAsync(() =>
            extension.NotifySessionClosedAsync("session1", null));

        Assert.Null(exception);
    }

    #endregion

    #region DisposeAsync Tests

    [Fact]
    public async Task DisposeAsync_MultipleTimes_DoesNotThrow()
    {
        var extension = CreateExtension();

        await extension.DisposeAsync();
        var exception = await Record.ExceptionAsync(async () => await extension.DisposeAsync());

        Assert.Null(exception);
    }

    [Fact]
    public async Task DisposeAsync_SetsDisposedFlag()
    {
        var extension = CreateExtension();

        await extension.DisposeAsync();

        var result = await extension.EnsureStartedAsync();
        Assert.False(result);
    }

    #endregion

    #region Helper Methods

    private ExtensionClass CreateExtension(ExtensionDefinition? definition = null)
    {
        return new ExtensionClass(
            definition ?? _definition,
            _config,
            _transportMock.Object,
            _snapshotManager,
            null,
            _loggerMock.Object);
    }

    private ExtensionClass CreateExtensionWithInvalidCommand()
    {
        var definition = new ExtensionDefinition
        {
            Id = "test-invalid",
            IsAvailable = true,
            Command = new ExtensionCommand
            {
                Type = "executable",
                Executable = "nonexistent_command_that_does_not_exist_12345"
            }
        };
        return CreateExtension(definition);
    }

    private static ExtensionDefinition CreateTestDefinition()
    {
        return new ExtensionDefinition
        {
            Id = "test-ext",
            IsAvailable = true,
            Command = new ExtensionCommand
            {
                Type = "executable",
                Executable = OperatingSystem.IsWindows() ? "cmd.exe" : "/bin/cat"
            },
            InputFormats = ["pdf"],
            SupportedDocumentTypes = ["word"],
            TransportModes = ["file"],
            Capabilities = new ExtensionCapabilities
            {
                SupportsHeartbeat = true
            }
        };
    }

    private static ExtensionMetadata CreateTestMetadata()
    {
        return new ExtensionMetadata
        {
            SessionId = $"test_session_{Guid.NewGuid():N}",
            SequenceNumber = 1,
            DocumentType = "word",
            OutputFormat = "pdf",
            MimeType = "application/pdf"
        };
    }

    #endregion
}
