using System.Text.Json;
using AsposeMcpServer.Core.Extension;
using Microsoft.Extensions.Logging;
using Moq;

namespace AsposeMcpServer.Tests.Core.Extension;

/// <summary>
///     Unit tests for the ExtensionManager class.
/// </summary>
public class ExtensionManagerTests : IAsyncDisposable
{
    private readonly Mock<ILoggerFactory> _loggerFactoryMock;
    private readonly Mock<ILogger<ExtensionManager>> _loggerMock;
    private readonly Mock<ILogger<SnapshotManager>> _snapshotLoggerMock;
    private readonly string _tempDir;
    private ExtensionConfig _config = null!;
    private SnapshotManager _snapshotManager = null!;

    public ExtensionManagerTests()
    {
        _loggerFactoryMock = new Mock<ILoggerFactory>();
        _loggerMock = new Mock<ILogger<ExtensionManager>>();
        _snapshotLoggerMock = new Mock<ILogger<SnapshotManager>>();

        _loggerFactoryMock.Setup(x => x.CreateLogger(It.IsAny<string>()))
            .Returns(Mock.Of<ILogger>());

        _tempDir = Path.Combine(Path.GetTempPath(), $"ExtensionManagerTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        InitializeConfig();
    }

    public async ValueTask DisposeAsync()
    {
        _snapshotManager.Dispose();

        try
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, true);
        }
        // ReSharper disable once EmptyGeneralCatchClause - Best-effort cleanup in dispose
        catch
        {
        }

        await Task.CompletedTask;
        GC.SuppressFinalize(this);
    }

    private void InitializeConfig(string? configPath = null)
    {
        _config = new ExtensionConfig
        {
            Enabled = true,
            ConfigPath = configPath,
            TempDirectory = _tempDir,
            MaxRestartAttempts = 3,
            HealthCheckIntervalSeconds = 30,
            HandshakeTimeoutSeconds = 1
        };
        _snapshotManager = new SnapshotManager(_config, _snapshotLoggerMock.Object);
    }

    #region DisposeAsync Tests

    [Fact]
    public async Task DisposeAsync_MultipleTimes_DoesNotThrow()
    {
        InitializeConfig();

        var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.DisposeAsync();
        var exception = await Record.ExceptionAsync(async () => await manager.DisposeAsync());

        Assert.Null(exception);
    }

    #endregion

    #region Helper Methods

    private string CreateTestConfigFile(ExtensionDefinition[] extensions)
    {
        var configFile = new ExtensionsConfigFile
        {
            Extensions = extensions.ToDictionary(e => e.Id, e => e)
        };
        var json = JsonSerializer.Serialize(configFile, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = true
        });

        var configPath = Path.Combine(_tempDir, $"extensions_{Guid.NewGuid():N}.json");
        File.WriteAllText(configPath, json);
        return configPath;
    }

    #endregion

    #region StartAsync Tests

    [Fact]
    public async Task StartAsync_WhenDisabled_CompletesWithoutError()
    {
        var config = new ExtensionConfig { Enabled = false, TempDirectory = _tempDir };
        var snapshotManager = new SnapshotManager(config, _snapshotLoggerMock.Object);

        await using var manager = new ExtensionManager(
            config,
            snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        var exception = await Record.ExceptionAsync(() => manager.StartAsync(CancellationToken.None));

        Assert.Null(exception);
        snapshotManager.Dispose();
    }

    [Fact]
    public async Task StartAsync_NoConfigPath_CompletesWithoutError()
    {
        InitializeConfig();

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        var exception = await Record.ExceptionAsync(() => manager.StartAsync(CancellationToken.None));

        Assert.Null(exception);
    }

    [Fact]
    public async Task StartAsync_ConfigFileNotFound_CompletesWithoutError()
    {
        var nonExistentPath = Path.Combine(_tempDir, "nonexistent.json");
        InitializeConfig(nonExistentPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        var exception = await Record.ExceptionAsync(() => manager.StartAsync(CancellationToken.None));

        Assert.Null(exception);
    }

    [Fact]
    public async Task StartAsync_ValidConfig_LoadsExtensions()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "test-ext",
                Command = new ExtensionCommand { Type = "executable", Executable = "cmd.exe" }
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var extensions = manager.ListExtensions().ToList();
        Assert.Single(extensions);
        Assert.Equal("test-ext", extensions[0].Id);
    }

    [Fact]
    public async Task StartAsync_InvalidJson_LogsWarningAndContinues()
    {
        var configPath = Path.Combine(_tempDir, "invalid.json");
        await File.WriteAllTextAsync(configPath, "{ invalid json }");
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        Assert.Empty(manager.ListExtensions());
    }

    [Fact]
    public void CreateTestConfigFile_DuplicateExtensionId_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() => CreateTestConfigFile(
        [
            new ExtensionDefinition { Id = "duplicate-id" },
            new ExtensionDefinition { Id = "duplicate-id" }
        ]));
    }

    [Fact]
    public async Task StartAsync_ExtensionIdWithColon_LogsWarningAndContinues()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition { Id = "invalid:id" }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        Assert.Empty(manager.ListExtensions());
    }

    [Fact]
    public async Task StartAsync_MissingExtensionId_SkipsExtension()
    {
        var configPath = Path.Combine(_tempDir, $"extensions_empty_id_{Guid.NewGuid():N}.json");
        await File.WriteAllTextAsync(configPath, """
                                                 {
                                                     "extensions": {
                                                         "": {
                                                             "command": { "type": "executable", "executable": "cmd.exe" }
                                                         }
                                                     }
                                                 }
                                                 """);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var extensions = manager.ListExtensions().ToList();
        Assert.Empty(extensions);
    }

    [Fact]
    public async Task StartAsync_WithCommand_AttemptsHandshakeAndMarksUnavailableOnFailure()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "test-ext",
                Command = new ExtensionCommand
                {
                    Type = "executable",
                    Executable = OperatingSystem.IsWindows()
                        ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "findstr.exe")
                        : "/bin/cat"
                },
                TransportModes = ["file"],
                Capabilities = new ExtensionCapabilities { IdleTimeoutMinutes = 0 }
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var timeout = DateTime.UtcNow.AddSeconds(30);
        ExtensionDefinition? extension = null;
        while (DateTime.UtcNow < timeout)
        {
            extension = manager.ListExtensions().FirstOrDefault();
            if (extension is { IsAvailable: false })
                break;
            await Task.Delay(100);
        }

        Assert.NotNull(extension);
        Assert.False(extension.IsAvailable, "Extension should be marked unavailable after handshake failure");
        Assert.True(
            extension.UnavailableReason?.Contains("Handshake") == true ||
            extension.UnavailableReason?.Contains("Initialization") == true,
            $"Expected 'Handshake' or 'Initialization' in UnavailableReason, got: {extension.UnavailableReason}");
    }

    [Fact]
    public async Task StartAsync_WithoutCommand_ExtensionRemainsAvailable()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "no-command-ext",
                SupportedDocumentTypes = ["word"],
                InputFormats = ["pdf"]
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var extensions = manager.ListExtensions().ToList();
        Assert.Single(extensions);
        Assert.True(extensions[0].IsAvailable);
    }

    #endregion

    #region StopAsync Tests

    [Fact]
    public async Task StopAsync_WhenDisabled_CompletesWithoutError()
    {
        var config = new ExtensionConfig { Enabled = false, TempDirectory = _tempDir };
        var snapshotManager = new SnapshotManager(config, _snapshotLoggerMock.Object);

        await using var manager = new ExtensionManager(
            config,
            snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        var exception = await Record.ExceptionAsync(() => manager.StopAsync(CancellationToken.None));

        Assert.Null(exception);
        snapshotManager.Dispose();
    }

    [Fact]
    public async Task StopAsync_AfterStart_CompletesSuccessfully()
    {
        InitializeConfig();

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);
        var exception = await Record.ExceptionAsync(() => manager.StopAsync(CancellationToken.None));

        Assert.Null(exception);
    }

    #endregion

    #region GetExtensionAsync Tests

    [Fact]
    public async Task GetExtensionAsync_UnknownId_ReturnsNull()
    {
        InitializeConfig();

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var result = await manager.GetExtensionAsync("unknown-extension");

        Assert.Null(result);
    }

    [Fact]
    public async Task GetExtensionAsync_UnavailableExtension_ReturnsNull()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "unavailable-ext",
                IsAvailable = false,
                UnavailableReason = "Executable not found"
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var result = await manager.GetExtensionAsync("unavailable-ext");

        Assert.Null(result);
    }

    #endregion

    #region GetRunningExtension Tests

    [Fact]
    public async Task GetRunningExtension_NotStarted_ReturnsNull()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "test-ext",
                Command = new ExtensionCommand { Type = "executable", Executable = "cmd.exe" }
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var result = manager.GetRunningExtension("test-ext");

        Assert.Null(result);
    }

    [Fact]
    public async Task GetRunningExtension_UnknownId_ReturnsNull()
    {
        InitializeConfig();

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var result = manager.GetRunningExtension("unknown");

        Assert.Null(result);
    }

    #endregion

    #region FindExtensionsForDocument Tests

    [Fact]
    public async Task FindExtensionsForDocument_MatchingExtension_ReturnsExtension()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "pdf-viewer",
                SupportedDocumentTypes = ["word", "excel"],
                InputFormats = ["pdf"]
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var results = manager.FindExtensionsForDocument("word", "pdf").ToList();

        Assert.Single(results);
        Assert.Equal("pdf-viewer", results[0].Id);
    }

    [Fact]
    public async Task FindExtensionsForDocument_NoMatch_ReturnsEmpty()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "pdf-viewer",
                SupportedDocumentTypes = ["word"],
                InputFormats = ["pdf"]
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var results = manager.FindExtensionsForDocument("powerpoint", "html").ToList();

        Assert.Empty(results);
    }

    [Fact]
    public async Task FindExtensionsForDocument_EmptyFilters_ReturnsAll()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "universal-ext",
                SupportedDocumentTypes = [],
                InputFormats = []
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var results = manager.FindExtensionsForDocument("anything", "anyformat").ToList();

        Assert.Single(results);
    }

    [Fact]
    public async Task FindExtensionsForDocument_UnavailableExtension_NotReturned()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "unavailable-ext",
                SupportedDocumentTypes = ["word"],
                InputFormats = ["pdf"],
                Command = new ExtensionCommand
                {
                    Type = "executable",
                    Executable = "/nonexistent/path/to/executable_that_does_not_exist"
                }
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var results = manager.FindExtensionsForDocument("word", "pdf").ToList();

        Assert.Empty(results);
    }

    #endregion

    #region ListExtensions Tests

    [Fact]
    public async Task ListExtensions_NoExtensions_ReturnsEmpty()
    {
        InitializeConfig();

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var results = manager.ListExtensions().ToList();

        Assert.Empty(results);
    }

    [Fact]
    public async Task ListExtensions_WithExtensions_ReturnsAll()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition { Id = "ext1" },
            new ExtensionDefinition { Id = "ext2" },
            new ExtensionDefinition { Id = "ext3" }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var results = manager.ListExtensions().ToList();

        Assert.Equal(3, results.Count);
    }

    #endregion

    #region GetExtensionStatuses Tests

    [Fact]
    public async Task GetExtensionStatuses_UnstartedExtension_ReturnsUnloadedState()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "test-ext",
                Command = new ExtensionCommand { Type = "executable", Executable = "cmd.exe" }
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var statuses = manager.GetExtensionStatuses();

        Assert.Single(statuses);
        Assert.True(statuses.ContainsKey("test-ext"));
        Assert.Equal(ExtensionState.Unloaded, statuses["test-ext"].State);
    }

    [Fact]
    public async Task GetExtensionStatuses_UnavailableExtension_IncludesReason()
    {
        var configPath = CreateTestConfigFile(
        [
            new ExtensionDefinition
            {
                Id = "unavailable-ext",
                Command = new ExtensionCommand
                {
                    Type = "executable",
                    Executable = "/nonexistent/path/to/executable_that_does_not_exist"
                }
            }
        ]);
        InitializeConfig(configPath);

        await using var manager = new ExtensionManager(
            _config,
            _snapshotManager,
            _loggerFactoryMock.Object,
            _loggerMock.Object);

        await manager.StartAsync(CancellationToken.None);

        var statuses = manager.GetExtensionStatuses();

        Assert.False(statuses["unavailable-ext"].IsAvailable);
        Assert.Contains("Executable not found", statuses["unavailable-ext"].UnavailableReason);
    }

    #endregion
}
