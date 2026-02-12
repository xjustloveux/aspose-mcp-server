using System.Collections.Concurrent;
using System.Text.Json;
using AsposeMcpServer.Core.Extension.Transport;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Manages extension lifecycle, discovery, and health monitoring.
///     Responsible for loading extension definitions, creating extension instances,
///     and performing periodic health checks.
/// </summary>
public class ExtensionManager : IHostedService, IAsyncDisposable
{
    /// <summary>
    ///     Valid transport modes supported by the system.
    /// </summary>
    /// <remarks>
    ///     <list type="bullet">
    ///         <item><c>mmap</c>: Memory-mapped file for high-performance, large data transfers (cross-platform)</item>
    ///         <item><c>stdin</c>: Standard input with length-prefix framing, cross-platform</item>
    ///         <item><c>file</c>: Temporary file transfer, most compatible but slower</item>
    ///     </list>
    /// </remarks>
    private static readonly string[] ValidTransportModes = ["mmap", "stdin", "file"];

    /// <summary>
    ///     Extension configuration containing settings for health checks and timeouts.
    /// </summary>
    private readonly ExtensionConfig _config;

    /// <summary>
    ///     Dictionary of extension definitions keyed by extension ID.
    /// </summary>
    private readonly ConcurrentDictionary<string, ExtensionDefinition> _definitions = new();

    /// <summary>
    ///     Tracks when extensions entered Error state for cooldown enforcement.
    ///     Key: extension ID, Value: UTC time when Error state was entered.
    /// </summary>
    private readonly ConcurrentDictionary<string, DateTime> _errorCooldowns = new();

    /// <summary>
    ///     Dictionary of active extension instances keyed by extension ID.
    ///     Uses Lazy to ensure only one Extension is created even in race conditions.
    /// </summary>
    private readonly ConcurrentDictionary<string, Lazy<Extension>> _extensions = new();

    /// <summary>
    ///     File transport instance for extensions using file-based communication.
    /// </summary>
    private readonly FileTransport _fileTransport;

    /// <summary>
    ///     Tracks extension IDs that are currently being initialized in background.
    /// </summary>
    private readonly ConcurrentDictionary<string, bool> _initializingExtensions = new();


    /// <summary>
    ///     Logger instance for diagnostic output.
    /// </summary>
    private readonly ILogger<ExtensionManager> _logger;

    /// <summary>
    ///     Logger factory for creating loggers for extension instances.
    /// </summary>
    private readonly ILoggerFactory _loggerFactory;

    /// <summary>
    ///     Memory-mapped file transport instance for high-performance communication.
    /// </summary>
    private readonly MmapTransport _mmapTransport;

    /// <summary>
    ///     Process cleanup manager for ensuring child processes are terminated on exit.
    /// </summary>
    private readonly ProcessCleanupManager _processCleanupManager;

    /// <summary>
    ///     Dictionary of active restart tasks keyed by extension ID.
    ///     Used to track and await restart operations during disposal.
    /// </summary>
    private readonly ConcurrentDictionary<string, Task> _restartTasks = new();

    /// <summary>
    ///     Snapshot manager for tracking pending snapshots.
    /// </summary>
    private readonly SnapshotManager _snapshotManager;

    /// <summary>
    ///     Stdin transport instance for binary stdin communication.
    /// </summary>
    private readonly StdinTransport _stdinTransport;

    /// <summary>
    ///     Whether this instance has been disposed.
    /// </summary>
    private bool _disposed;

    /// <summary>
    ///     Cancellation token source for the health check loop.
    /// </summary>
    private CancellationTokenSource? _healthCheckCts;

    /// <summary>
    ///     The background health check task.
    /// </summary>
    private Task? _healthCheckTask;

    /// <summary>
    ///     Background initialization task for extensions.
    /// </summary>
    private Task? _initializationTask;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExtensionManager" /> class.
    /// </summary>
    /// <param name="config">Extension configuration.</param>
    /// <param name="snapshotManager">Snapshot manager.</param>
    /// <param name="loggerFactory">Logger factory.</param>
    /// <param name="logger">Logger instance.</param>
    public ExtensionManager(
        ExtensionConfig config,
        SnapshotManager snapshotManager,
        ILoggerFactory loggerFactory,
        ILogger<ExtensionManager> logger)
    {
        _config = config;
        _snapshotManager = snapshotManager;
        _loggerFactory = loggerFactory;
        _logger = logger;

        FileTransport.CleanupOrphanedDirectories(config.TempDirectory, logger);

        FileTransport? fileTransport = null;
        MmapTransport? mmapTransport = null;
        ProcessCleanupManager? processCleanupManager = null;

        try
        {
            var snapshotDir = FileTransport.GenerateDirectoryWithPid(config.TempDirectory);
            fileTransport = new FileTransport(
                snapshotDir,
                loggerFactory.CreateLogger<FileTransport>(),
                config.MaxSnapshotSizeBytes,
                config.MinFreeDiskSpaceBytes);
            var stdinTransport = new StdinTransport(
                loggerFactory.CreateLogger<StdinTransport>(),
                maxDataSize: config.MaxSnapshotSizeBytes);
            mmapTransport = new MmapTransport(
                loggerFactory.CreateLogger<MmapTransport>(),
                config.MaxSnapshotSizeBytes,
                config.TempDirectory);
            processCleanupManager = new ProcessCleanupManager(logger);

            _fileTransport = fileTransport;
            _stdinTransport = stdinTransport;
            _mmapTransport = mmapTransport;
            _processCleanupManager = processCleanupManager;
        }
        catch
        {
            fileTransport?.Dispose();
            mmapTransport?.Dispose();
            processCleanupManager?.Dispose();
            throw;
        }
    }

    /// <summary>
    ///     Gets the process cleanup manager for registering extension processes.
    /// </summary>
    internal ProcessCleanupManager ProcessCleanupManager => _processCleanupManager;

    /// <inheritdoc />
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        if (_healthCheckCts != null)
            await _healthCheckCts.CancelAsync();

        if (_healthCheckTask != null)
            try
            {
                await _healthCheckTask.WaitAsync(TimeSpan.FromSeconds(5));
            }
            catch (OperationCanceledException)
            {
            }
            catch (TimeoutException)
            {
                _logger.LogDebug("Health check task did not complete within timeout during disposal");
            }

        if (_initializationTask != null)
            try
            {
                await _initializationTask.WaitAsync(TimeSpan.FromSeconds(5));
            }
            catch (OperationCanceledException)
            {
            }
            catch (TimeoutException)
            {
                _logger.LogDebug("Background initialization task did not complete within timeout during disposal");
            }

        _healthCheckCts?.Dispose();

        var restartTasks = _restartTasks.Values.ToArray();
        if (restartTasks.Length > 0)
        {
            _logger.LogDebug("Waiting for {Count} pending restart task(s) to complete", restartTasks.Length);
            try
            {
                await Task.WhenAll(restartTasks).WaitAsync(TimeSpan.FromSeconds(10));
            }
            catch (TimeoutException)
            {
                _logger.LogWarning("Timed out waiting for restart tasks during disposal");
            }
            catch (Exception ex)
            {
                _logger.LogDebug(ex, "Error waiting for restart tasks during disposal");
            }
        }

        foreach (var lazy in _extensions.Values)
            if (lazy.IsValueCreated)
                await lazy.Value.DisposeAsync();

        _extensions.Clear();
        _restartTasks.Clear();
        _fileTransport.Dispose();
        _mmapTransport.Dispose();
        _processCleanupManager.Dispose();

        _disposed = true;
        GC.SuppressFinalize(this);
    }

    /// <inheritdoc />
    public async Task StartAsync(CancellationToken cancellationToken)
    {
        if (!_config.Enabled)
        {
            _logger.LogInformation("Extension system is disabled");
            return;
        }

        try
        {
            await LoadExtensionDefinitionsAsync(cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to load extension definitions, extension system will be disabled");
            return;
        }

        _healthCheckCts = new CancellationTokenSource();
        _healthCheckTask = RunHealthCheckLoopAsync(_healthCheckCts.Token);

        var extensionsToInit = _definitions.Values
            .Where(d => d is { IsAvailable: true, HasValidCommand: true })
            .ToList();

        if (extensionsToInit.Count > 0)
            _initializationTask = Task.Run(
                () => InitializeExtensionsInBackgroundAsync(extensionsToInit, _healthCheckCts.Token),
                _healthCheckCts.Token);

        _logger.LogInformation(
            "ExtensionManager started with {Count} extension(s), initialization proceeding in background",
            _definitions.Count);
    }

    /// <inheritdoc />
    public async Task StopAsync(CancellationToken cancellationToken)
    {
        if (!_config.Enabled)
            return;

        if (_healthCheckCts != null)
        {
            await _healthCheckCts.CancelAsync();
            if (_healthCheckTask != null)
                try
                {
                    await _healthCheckTask.WaitAsync(cancellationToken);
                }
                catch (OperationCanceledException)
                {
                }
        }

        var stopTasks = _extensions.Values
            .Where(lazy => lazy.IsValueCreated)
            .Select(lazy => lazy.Value.StopAsync())
            .ToList();
        await Task.WhenAll(stopTasks);

        _logger.LogInformation("ExtensionManager stopped");
    }

    /// <summary>
    ///     Event raised when an extension enters Error state.
    ///     Used by ExtensionSessionBridge to clean up bindings.
    /// </summary>
    public event Action<string>? ExtensionError;

    /// <summary>
    ///     Gets an extension by ID, ensuring it is started.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    /// <returns>The extension instance, or null if not found.</returns>
    public async Task<Extension?> GetExtensionAsync(string extensionId)
    {
        if (!_definitions.TryGetValue(extensionId, out var definition))
            return null;

        if (!definition.IsAvailable)
            return null;

        var lazyExtension = _extensions.GetOrAdd(
            extensionId,
            _ => new Lazy<Extension>(() => CreateExtension(definition)));

        var extension = lazyExtension.Value;

        if (await extension.EnsureStartedAsync())
            return extension;

        return null;
    }

    /// <summary>
    ///     Gets an extension by ID only if it is already running (Idle or Busy state).
    ///     Does not start the extension if it is not running.
    /// </summary>
    /// <param name="extensionId">Extension identifier.</param>
    /// <returns>The extension instance if running, or null if not found or not running.</returns>
    public Extension? GetRunningExtension(string extensionId)
    {
        if (!_extensions.TryGetValue(extensionId, out var lazy))
            return null;

        if (!lazy.IsValueCreated)
            return null;

        var extension = lazy.Value;
        var state = extension.State;

        if (state == ExtensionState.Idle || state == ExtensionState.Busy)
            return extension;

        return null;
    }

    /// <summary>
    ///     Finds extensions that can handle the given document type and output format.
    /// </summary>
    /// <param name="documentType">Document type (e.g., "word", "excel").</param>
    /// <param name="outputFormat">Output format (e.g., "pdf", "html").</param>
    /// <returns>List of matching extension definitions.</returns>
    public IEnumerable<ExtensionDefinition> FindExtensionsForDocument(string documentType, string outputFormat)
    {
        return _definitions.Values
            .Where(d => d.IsAvailable)
            .Where(d =>
            {
                var supportsDocType = d.SupportedDocumentTypes.Count == 0 ||
                                      d.SupportedDocumentTypes.Contains(documentType, StringComparer.OrdinalIgnoreCase);
                var supportsFormat = d.InputFormats.Count == 0 ||
                                     d.InputFormats.Contains(outputFormat, StringComparer.OrdinalIgnoreCase);
                return supportsDocType && supportsFormat;
            });
    }

    /// <summary>
    ///     Lists all registered extension definitions.
    /// </summary>
    /// <returns>All extension definitions.</returns>
    public IEnumerable<ExtensionDefinition> ListExtensions()
    {
        return _definitions.Values;
    }

    /// <summary>
    ///     Gets the current status of all extensions.
    /// </summary>
    /// <returns>Dictionary of extension ID to state.</returns>
    public Dictionary<string, ExtensionStatusInfo> GetExtensionStatuses()
    {
        var result = new Dictionary<string, ExtensionStatusInfo>();

        foreach (var definition in _definitions.Values)
        {
            var status = new ExtensionStatusInfo
            {
                Id = definition.Id,
                Name = definition.DisplayName,
                Version = definition.DisplayVersion,
                Title = definition.DisplayTitle,
                Description = definition.DisplayDescription,
                Author = definition.DisplayAuthor,
                WebsiteUrl = definition.DisplayWebsiteUrl,
                IsAvailable = definition.IsAvailable,
                UnavailableReason = definition.UnavailableReason
            };

            if (_extensions.TryGetValue(definition.Id, out var lazy) && lazy.IsValueCreated)
            {
                var ext = lazy.Value;
                status.State = ext.State;
                status.LastActivity = ext.LastActivity;
                status.RestartCount = ext.RestartCount;
                status.IsInitializing = ext.State == ExtensionState.Initializing;
            }
            else
            {
                status.State = ExtensionState.Unloaded;
                status.IsInitializing = _initializingExtensions.ContainsKey(definition.Id);
            }

            result[definition.Id] = status;
        }

        return result;
    }

    /// <summary>
    ///     Loads extension definitions from the configuration file.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>A task representing the loading operation.</returns>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when the configuration file contains invalid JSON or duplicate extension IDs.
    /// </exception>
    private async Task LoadExtensionDefinitionsAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrEmpty(_config.ConfigPath))
        {
            _logger.LogInformation("No extensions config path specified, skipping extension loading");
            return;
        }

        if (!File.Exists(_config.ConfigPath))
        {
            _logger.LogWarning("Extensions config file not found: {Path}", _config.ConfigPath);
            return;
        }

        try
        {
            var json = await File.ReadAllTextAsync(_config.ConfigPath, cancellationToken);
            var configFile = JsonSerializer.Deserialize<ExtensionsConfigFile>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (configFile?.Extensions == null || configFile.Extensions.Count == 0)
            {
                _logger.LogInformation("No extensions defined in config file");
                return;
            }

            var configDir = Path.GetDirectoryName(_config.ConfigPath) ?? string.Empty;

            foreach (var (id, definition) in configFile.Extensions)
            {
                if (string.IsNullOrWhiteSpace(id))
                {
                    _logger.LogWarning("Skipping extension with empty ID");
                    continue;
                }

                if (id.Contains(':'))
                {
                    _logger.LogError("Extension ID cannot contain colon character: {Id}", id);
                    throw new InvalidOperationException(
                        $"Extension ID cannot contain colon character: {id}");
                }

                definition.Id = id;

                ResolveRelativePaths(definition, configDir);
                ValidateDefinition(definition);

                _definitions[id] = definition;
                _logger.LogDebug("Loaded extension definition: {Id}", id);
            }
        }
        catch (JsonException ex)
        {
            _logger.LogError(ex, "Failed to parse extensions config file: {Path}", _config.ConfigPath);
            throw new InvalidOperationException($"Invalid JSON in extensions config file: {_config.ConfigPath}", ex);
        }
    }

    /// <summary>
    ///     Initializes extensions in background without blocking MCP startup.
    /// </summary>
    /// <param name="extensions">List of extension definitions to initialize.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>A task representing the background initialization.</returns>
    private async Task InitializeExtensionsInBackgroundAsync(
        IReadOnlyList<ExtensionDefinition> extensions,
        CancellationToken cancellationToken)
    {
        var successCount = 0;
        var failCount = 0;

        foreach (var definition in extensions)
        {
            if (cancellationToken.IsCancellationRequested)
                break;

            try
            {
                _initializingExtensions[definition.Id] = true;
                var success = await InitializeSingleExtensionAsync(definition, cancellationToken);

                if (success)
                    successCount++;
                else
                    failCount++;
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Unexpected error initializing extension {Id}", definition.Id);
                failCount++;
            }
            finally
            {
                _initializingExtensions.TryRemove(definition.Id, out _);
            }
        }

        _logger.LogInformation(
            "Background extension initialization completed: {Success} succeeded, {Failed} failed",
            successCount, failCount);
    }

    /// <summary>
    ///     Initializes a single extension: starts process and performs handshake.
    /// </summary>
    /// <param name="definition">The extension definition.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>True if initialization succeeded, false otherwise.</returns>
    private async Task<bool> InitializeSingleExtensionAsync(
        ExtensionDefinition definition,
        CancellationToken cancellationToken)
    {
        try
        {
            var extension = await GetExtensionAsync(definition.Id);
            if (extension == null)
            {
                _logger.LogWarning("Failed to create extension instance for {Id}", definition.Id);
                return false;
            }

            await extension.PerformHandshakeAsync(cancellationToken);

            _logger.LogInformation(
                "Extension {Id} initialized: {Name} v{Version}",
                definition.Id,
                definition.DisplayName,
                definition.DisplayVersion);

            return true;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (OperationCanceledException ex)
        {
            _logger.LogWarning("Handshake timeout for extension {Id}", definition.Id);
            definition.IsAvailable = false;
            definition.UnavailableReason = $"Handshake timeout: {ex.Message}";
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to initialize extension {Id}", definition.Id);
            definition.IsAvailable = false;
            definition.UnavailableReason = $"Initialization failed: {ex.Message}";
            return false;
        }
    }

    /// <summary>
    ///     Resolves relative paths in the extension definition to absolute paths.
    ///     Validates that resolved paths do not escape the configuration directory to prevent path traversal attacks.
    /// </summary>
    /// <param name="definition">The extension definition to update.</param>
    /// <param name="configDir">The directory containing the configuration file.</param>
    /// <exception cref="InvalidOperationException">Thrown when a path attempts to escape the config directory.</exception>
    private static void ResolveRelativePaths(ExtensionDefinition definition, string configDir)
    {
        var command = definition.Command;
        var normalizedConfigDir = Path.GetFullPath(configDir)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (!string.IsNullOrEmpty(command.Executable) && !Path.IsPathRooted(command.Executable))
        {
            var resolvedPath = Path.GetFullPath(Path.Combine(configDir, command.Executable));
            ValidatePathWithinDirectory(resolvedPath, normalizedConfigDir, "Executable", definition.Id);
            command.Executable = resolvedPath;
        }

        if (!string.IsNullOrEmpty(command.WorkingDirectory) && !Path.IsPathRooted(command.WorkingDirectory))
        {
            var resolvedPath = Path.GetFullPath(Path.Combine(configDir, command.WorkingDirectory));
            ValidatePathWithinDirectory(resolvedPath, normalizedConfigDir, "WorkingDirectory", definition.Id);
            command.WorkingDirectory = resolvedPath;
        }
    }

    /// <summary>
    ///     Validates that a resolved path is within the allowed base directory.
    ///     Prevents path traversal attacks using ".." sequences.
    /// </summary>
    /// <param name="resolvedPath">The fully resolved path to validate.</param>
    /// <param name="baseDirectory">The base directory that paths must be within.</param>
    /// <param name="pathType">Type of path being validated (for error messages).</param>
    /// <param name="extensionId">Extension ID (for error messages).</param>
    /// <exception cref="InvalidOperationException">Thrown when the path escapes the base directory.</exception>
    private static void ValidatePathWithinDirectory(string resolvedPath, string baseDirectory, string pathType,
        string extensionId)
    {
        var normalizedResolved = Path.GetFullPath(resolvedPath)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

        if (!normalizedResolved.StartsWith(baseDirectory + Path.DirectorySeparatorChar,
                StringComparison.OrdinalIgnoreCase) &&
            !normalizedResolved.Equals(baseDirectory, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                $"Extension '{extensionId}' has invalid {pathType}: path traversal detected. " +
                $"Path '{resolvedPath}' escapes configuration directory '{baseDirectory}'");
    }

    /// <summary>
    ///     Validates an extension definition and marks it unavailable if validation fails.
    /// </summary>
    /// <param name="definition">The extension definition to validate.</param>
    private void ValidateDefinition(ExtensionDefinition definition)
    {
        var command = definition.Command;

        if (!definition.HasValidCommand)
        {
            _logger.LogDebug(
                "Extension {Id} has no executable configured, will be used for metadata only",
                definition.Id);
            return;
        }

        if (command.Type is "executable" or "custom")
        {
            if (!File.Exists(command.Executable))
            {
                definition.IsAvailable = false;
                definition.UnavailableReason = $"Executable not found: {command.Executable}";
                _logger.LogWarning(
                    "Extension {Id} marked unavailable: {Reason}",
                    definition.Id, definition.UnavailableReason);
                return;
            }
        }
        else if (command.Type is "node" or "python" or "dotnet")
        {
            if (!File.Exists(command.Executable))
            {
                definition.IsAvailable = false;
                definition.UnavailableReason = $"Script not found: {command.Executable}";
                _logger.LogWarning(
                    "Extension {Id} marked unavailable: {Reason}",
                    definition.Id, definition.UnavailableReason);
                return;
            }
        }

        var invalidModes = definition.TransportModes
            .Where(m => !ValidTransportModes.Contains(m, StringComparer.OrdinalIgnoreCase))
            .ToList();

        if (invalidModes.Count > 0)
        {
            _logger.LogWarning(
                "Extension {Id} has unsupported transport modes: {Modes}. Valid modes are: {ValidModes}",
                definition.Id, string.Join(", ", invalidModes), string.Join(", ", ValidTransportModes));

            definition.TransportModes = definition.TransportModes
                .Where(m => ValidTransportModes.Contains(m, StringComparer.OrdinalIgnoreCase))
                .ToList();

            if (definition.TransportModes.Count == 0)
                definition.TransportModes = ["file"];
        }

        if (!string.IsNullOrEmpty(definition.PreferredTransportMode) &&
            !ValidTransportModes.Contains(definition.PreferredTransportMode, StringComparer.OrdinalIgnoreCase))
        {
            _logger.LogWarning(
                "Extension {Id} has invalid preferred transport mode: {Mode}. Falling back to first available.",
                definition.Id, definition.PreferredTransportMode);
            definition.PreferredTransportMode = null;
        }
    }

    /// <summary>
    ///     Creates a new extension instance from a definition.
    /// </summary>
    /// <param name="definition">The extension definition.</param>
    /// <returns>A new <see cref="Extension" /> instance.</returns>
    private Extension CreateExtension(ExtensionDefinition definition)
    {
        var transport = SelectTransport(definition);
        var logger = _loggerFactory.CreateLogger<Extension>();

        var constraintWarnings = definition.ValidateCapabilityConstraints(_config);
        foreach (var warning in constraintWarnings)
            _logger.LogWarning("Extension {ExtensionId}: {Warning}", definition.Id, warning);

        var extension = new Extension(
            definition,
            _config,
            transport,
            _snapshotManager,
            _processCleanupManager,
            logger);

        extension.StateChanged += OnExtensionStateChanged;

        return extension;
    }

    /// <summary>
    ///     Selects the appropriate transport for an extension based on its configuration.
    /// </summary>
    /// <param name="definition">The extension definition.</param>
    /// <returns>The selected transport instance.</returns>
    /// <remarks>
    ///     Transport selection priority:
    ///     <list type="number">
    ///         <item><see cref="ExtensionDefinition.PreferredTransportMode" /> if specified</item>
    ///         <item>First entry in <see cref="ExtensionDefinition.TransportModes" /> if not empty</item>
    ///         <item><see cref="ExtensionConfig.DefaultTransportMode" /> as fallback</item>
    ///     </list>
    ///     The selected mode must also be present in the extension's supported transport modes.
    /// </remarks>
    private IExtensionTransport SelectTransport(ExtensionDefinition definition)
    {
        var preferredMode = definition.PreferredTransportMode ??
                            (definition.TransportModes.Count > 0 ? definition.TransportModes[0] : null) ??
                            _config.DefaultTransportMode;

        IExtensionTransport selected = preferredMode.ToLowerInvariant() switch
        {
            "mmap" when definition.TransportModes.Contains("mmap", StringComparer.OrdinalIgnoreCase) => _mmapTransport,
            "stdin" when definition.TransportModes.Contains("stdin", StringComparer.OrdinalIgnoreCase) =>
                _stdinTransport,
            _ => _fileTransport
        };

        _logger.LogTrace(
            "Selected transport for extension {ExtensionId}: {TransportMode} (preferred={Preferred}, supported=[{Supported}])",
            definition.Id, selected.Mode, preferredMode, string.Join(", ", definition.TransportModes));

        return selected;
    }

    /// <summary>
    ///     Event handler called when an extension's state changes.
    /// </summary>
    /// <param name="sender">The extension that raised the event.</param>
    /// <param name="state">The new state.</param>
    private void OnExtensionStateChanged(object? sender, ExtensionState state)
    {
        if (sender is not Extension extension)
            return;

        switch (state)
        {
            case ExtensionState.Crashed:
                TrackRestartTask(extension);
                break;

            case ExtensionState.Error:
                _ = HandleExtensionErrorAsync(extension);
                break;
        }
    }

    /// <summary>
    ///     Attempts to recover an extension from Error state.
    ///     Enforces a cooldown period before allowing recovery.
    ///     Prevents concurrent recovery/restart attempts.
    /// </summary>
    /// <param name="extensionId">The extension ID to recover.</param>
    /// <returns>True if recovery was successful.</returns>
    public async Task<(bool Success, string? Error)> TryRecoverExtensionAsync(string extensionId)
    {
        _logger.LogDebug("Attempting to recover extension {ExtensionId}", extensionId);

        if (!_definitions.TryGetValue(extensionId, out var definition))
        {
            _logger.LogTrace("Recovery failed: extension {ExtensionId} not found in definitions", extensionId);
            return (false, $"Extension not found: {extensionId}");
        }

        if (!definition.IsAvailable)
        {
            _logger.LogTrace(
                "Recovery failed: extension {ExtensionId} is not available ({Reason})",
                extensionId, definition.UnavailableReason);
            return (false, $"Extension is not available: {definition.UnavailableReason}");
        }

        if (_restartTasks.TryGetValue(extensionId, out var restartTask) && !restartTask.IsCompleted)
        {
            _logger.LogTrace(
                "Recovery failed: extension {ExtensionId} has pending restart task",
                extensionId);
            return (false, "Extension restart is already in progress. Please wait for it to complete.");
        }

        if (_errorCooldowns.TryGetValue(extensionId, out var errorTime))
        {
            var elapsed = DateTime.UtcNow - errorTime;
            var cooldownSeconds = _config.ErrorRecoveryCooldownSeconds;
            if (elapsed.TotalSeconds < cooldownSeconds)
            {
                var remaining = cooldownSeconds - (int)elapsed.TotalSeconds;
                _logger.LogTrace(
                    "Recovery failed: extension {ExtensionId} is in cooldown ({Elapsed:F1}s / {Total}s)",
                    extensionId, elapsed.TotalSeconds, cooldownSeconds);
                return (false, $"Extension is in cooldown period. Please wait {remaining} seconds before retrying.");
            }

            _logger.LogTrace(
                "Extension {ExtensionId} cooldown period elapsed ({Elapsed:F1}s >= {Total}s)",
                extensionId, elapsed.TotalSeconds, cooldownSeconds);
        }

        if (_extensions.TryGetValue(extensionId, out var lazy) && lazy.IsValueCreated)
        {
            var extension = lazy.Value;
            if (extension.State == ExtensionState.Error)
            {
                if (await extension.TryRecoverFromErrorAsync())
                {
                    _errorCooldowns.TryRemove(extensionId, out _);
                    _logger.LogInformation("Extension {ExtensionId} recovered successfully", extensionId);
                    return (true, null);
                }

                return (false, "Recovery attempt failed. Extension may have underlying issues.");
            }

            if (extension.State == ExtensionState.Idle || extension.State == ExtensionState.Busy)
            {
                _logger.LogTrace(
                    "Recovery not needed: extension {ExtensionId} is already running (state={State})",
                    extensionId, extension.State);
                return (false, "Extension is already running and does not need recovery.");
            }

            _logger.LogTrace(
                "Recovery failed: extension {ExtensionId} is in {State} state, expected Error",
                extensionId, extension.State);
            return (false, $"Extension is in {extension.State} state, not Error state.");
        }

        _errorCooldowns.TryRemove(extensionId, out _);
        var newExtension = await GetExtensionAsync(extensionId);
        if (newExtension != null)
        {
            _logger.LogInformation("Extension {ExtensionId} recovered by creating new instance", extensionId);
            return (true, null);
        }

        return (false, "Failed to create new extension instance.");
    }

    /// <summary>
    ///     Handles an extension entering Error state by cleaning up resources.
    /// </summary>
    /// <param name="extension">The extension in error state.</param>
    /// <returns>A task representing the cleanup operation.</returns>
    private async Task HandleExtensionErrorAsync(Extension extension)
    {
        var extensionId = extension.Definition.Id;

        _errorCooldowns[extensionId] = DateTime.UtcNow;

        _logger.LogWarning(
            "Extension {ExtensionId} entered Error state, cleaning up resources. " +
            "Recovery will be available after {CooldownSeconds} seconds.",
            extensionId, _config.ErrorRecoveryCooldownSeconds);

        ExtensionError?.Invoke(extensionId);

        if (_extensions.TryRemove(extensionId, out var lazy))
            if (lazy.IsValueCreated)
            {
                extension.StateChanged -= OnExtensionStateChanged;
                await extension.DisposeAsync();
            }
    }

    /// <summary>
    ///     Creates and tracks a restart task for a crashed extension.
    ///     Ensures the task is removed from tracking when complete.
    ///     Prevents recursive restart attempts by checking for existing tasks.
    /// </summary>
    /// <param name="extension">The crashed extension.</param>
    /// <remarks>
    ///     <para>
    ///         This method is called from the StateChanged event handler when an extension
    ///         enters the Crashed state. The restart is performed asynchronously to avoid
    ///         blocking the event handler.
    ///     </para>
    ///     <para>
    ///         Task tracking serves two purposes:
    ///         <list type="number">
    ///             <item>Prevent concurrent restart attempts for the same extension</item>
    ///             <item>Allow graceful shutdown to wait for pending restarts during disposal</item>
    ///         </list>
    ///     </para>
    ///     <para>
    ///         If a restart task is already in progress (not completed), new crash events
    ///         are ignored. The existing task will handle the restart.
    ///     </para>
    /// </remarks>
    private void TrackRestartTask(Extension extension)
    {
        var extensionId = extension.Definition.Id;

        if (_restartTasks.TryGetValue(extensionId, out var existingTask) && !existingTask.IsCompleted)
        {
            _logger.LogDebug(
                "Restart task already in progress for extension {Id}, skipping",
                extensionId);
            return;
        }

        var task = Task.Run(async () =>
        {
            try
            {
                await HandleExtensionCrashedAsync(extension);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex,
                    "Unhandled exception while handling crashed extension {Id}",
                    extensionId);
            }
            finally
            {
                _restartTasks.TryRemove(extensionId, out _);
            }
        });

        _restartTasks[extensionId] = task;
    }

    /// <summary>
    ///     Handles a crashed extension by attempting to restart it.
    /// </summary>
    /// <param name="extension">The crashed extension.</param>
    /// <returns>A task representing the restart operation.</returns>
    private async Task HandleExtensionCrashedAsync(Extension extension)
    {
        _logger.LogWarning("Extension {Id} crashed, attempting restart", extension.Definition.Id);

        if (!await extension.TryRestartAsync())
            _logger.LogError("Failed to restart extension {Id}", extension.Definition.Id);
    }

    /// <summary>
    ///     Runs the background health check loop.
    /// </summary>
    /// <param name="cancellationToken">Token to cancel the loop.</param>
    /// <returns>A task representing the health check loop.</returns>
    private async Task RunHealthCheckLoopAsync(CancellationToken cancellationToken)
    {
        var interval = TimeSpan.FromSeconds(_config.HealthCheckIntervalSeconds);

        while (!cancellationToken.IsCancellationRequested)
            try
            {
                await Task.Delay(interval, cancellationToken);
                await PerformHealthChecksAsync(cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in health check loop");
            }
    }

    /// <summary>
    ///     Performs health checks on all active extensions in parallel.
    ///     Unloads idle extensions and sends heartbeats to running ones.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>A task representing the health check operation.</returns>
    private async Task PerformHealthChecksAsync(CancellationToken cancellationToken)
    {
        var now = DateTime.UtcNow;
        var extensionsToCheck = new List<Extension>();

        foreach (var lazy in _extensions.Values)
        {
            if (!lazy.IsValueCreated)
                continue;

            var ext = lazy.Value;
            if (ext.State is not (ExtensionState.Unloaded or ExtensionState.Error or ExtensionState.Stopping))
                extensionsToCheck.Add(ext);
        }

        if (extensionsToCheck.Count == 0)
            return;

        var tasks = extensionsToCheck.Select(ext =>
            PerformSingleHealthCheckAsync(ext, now, cancellationToken));

        await Task.WhenAll(tasks);
    }

    /// <summary>
    ///     Performs health check on a single extension with timeout protection.
    /// </summary>
    /// <param name="extension">The extension to check.</param>
    /// <param name="now">Current UTC time.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>A task representing the health check operation.</returns>
    private async Task PerformSingleHealthCheckAsync(
        Extension extension,
        DateTime now,
        CancellationToken cancellationToken)
    {
        try
        {
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(_config.HealthCheckIntervalSeconds / 2.0));

            var effectiveIdleMinutes = extension.Definition.GetEffectiveIdleTimeoutMinutes(_config);

            if (effectiveIdleMinutes > 0)
            {
                var idleTimeout = TimeSpan.FromMinutes(effectiveIdleMinutes);
                if (now - extension.LastActivity > idleTimeout)
                {
                    _logger.LogInformation(
                        "Unloading idle extension {Id} after {Minutes} minutes of inactivity",
                        extension.Definition.Id, effectiveIdleMinutes);
                    await extension.StopAsync(true);
                    return;
                }
            }

            if (extension.State == ExtensionState.Idle)
                await extension.SendHeartbeatAsync(timeoutCts.Token);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            _logger.LogWarning(
                "Health check for extension {Id} timed out",
                extension.Definition.Id);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex,
                "Error during health check for extension {Id}",
                extension.Definition.Id);
        }
    }
}
