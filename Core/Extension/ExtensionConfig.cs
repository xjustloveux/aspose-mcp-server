using System.Diagnostics.CodeAnalysis;
using AsposeMcpServer.Core.Session;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Configuration for the extension system.
/// </summary>
public class ExtensionConfig
{
    /// <summary>
    ///     Whether extension system is enabled.
    /// </summary>
    public bool Enabled { get; set; }

    /// <summary>
    ///     Path to extensions.json configuration file.
    /// </summary>
    public string? ConfigPath { get; set; }

    /// <summary>
    ///     Temporary directory for snapshot files.
    /// </summary>
    [SuppressMessage("SonarAnalyzer.CSharp", "S5443",
        Justification = "System temp directory is validated in Validate() to prevent use of system directories")]
    public string TempDirectory { get; set; } = Path.GetTempPath();

    /// <summary>
    ///     Default transport mode if extension doesn't specify preference. Default is "stdin".
    /// </summary>
    public string DefaultTransportMode { get; set; } = "stdin";

    /// <summary>
    ///     Health check interval in seconds. Default is 30.
    /// </summary>
    public int HealthCheckIntervalSeconds { get; set; } = 30;

    /// <summary>
    ///     Maximum number of restart attempts for crashed extensions. Default is 3.
    /// </summary>
    public int MaxRestartAttempts { get; set; } = 3;

    /// <summary>
    ///     Cooldown period in seconds between restart attempts. Default is 5.
    /// </summary>
    public int RestartCooldownSeconds { get; set; } = 5;

    /// <summary>
    ///     Cooldown period in seconds before allowing recovery from Error state. Default is 60.
    ///     This prevents rapid retry loops when an extension has fundamental issues.
    /// </summary>
    public int ErrorRecoveryCooldownSeconds { get; set; } = 60;

    /// <summary>
    ///     Timeout in seconds for graceful shutdown of extensions. Default is 5.
    /// </summary>
    public int GracefulShutdownTimeoutSeconds { get; set; } = 5;

    /// <summary>
    ///     Timeout in seconds for extension handshake during initialization. Default is 30.
    ///     Extensions must respond to the "initialize" message within this time.
    /// </summary>
    public int HandshakeTimeoutSeconds { get; set; } = 30;

    /// <summary>
    ///     Timeout in milliseconds for stdin write operations. Default is 5000.
    ///     Prevents indefinite blocking if extension's stdin buffer is full.
    /// </summary>
    public int StdinWriteTimeoutMs { get; set; } = 5000;

    /// <summary>
    ///     Maximum number of consecutive snapshot send failures before marking extension as crashed.
    ///     Default is 10. Set to 0 to disable this check.
    /// </summary>
    public int MaxConsecutiveSendFailures { get; set; } = 10;

    /// <summary>
    ///     Timeout in seconds for the retry loop processing pending snapshots.
    ///     Default is 60. Prevents unbounded retry loops when many bindings are pending.
    /// </summary>
    public int RetryLoopTimeoutSeconds { get; set; } = 60;

    /// <summary>
    ///     Time-to-live in seconds for conversion cache entries. Default is 5.
    /// </summary>
    public int ConversionCacheTtlSeconds { get; set; } = 5;

    /// <summary>
    ///     Maximum number of entries in the conversion cache. Default is 100.
    /// </summary>
    public int MaxConversionCacheSize { get; set; } = 100;

    /// <summary>
    ///     Maximum number of consecutive conversion failures before entering backoff. Default is 5.
    /// </summary>
    public int MaxConversionFailures { get; set; } = 5;

    /// <summary>
    ///     Backoff duration in seconds after max conversion failures. Default is 60.
    /// </summary>
    public int FailureBackoffSeconds { get; set; } = 60;

    /// <summary>
    ///     Maximum snapshot size in bytes. Default is 100 MB.
    ///     Snapshots larger than this will be rejected to prevent memory exhaustion.
    /// </summary>
    public long MaxSnapshotSizeBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>
    ///     Minimum free disk space in bytes required before writing snapshot files. Default is 500 MB.
    ///     File transport will fail if available disk space is below this threshold.
    /// </summary>
    public long MinFreeDiskSpaceBytes { get; set; } = 500 * 1024 * 1024;

    /// <summary>
    ///     Frame interval configuration with default and constraints.
    ///     Extensions can override within [Floor, Ceiling] range.
    /// </summary>
    public ConstrainedInt FrameIntervalMs { get; set; } = new(
        100, 10, 5000);

    /// <summary>
    ///     Snapshot TTL configuration with default and constraints.
    ///     Extensions can override within [Floor, Ceiling] range.
    /// </summary>
    public ConstrainedInt SnapshotTtlSeconds { get; set; } = new(
        30, 5, 300);

    /// <summary>
    ///     Max missed heartbeats configuration with default and constraints.
    ///     Extensions can override within [Floor, Ceiling] range.
    /// </summary>
    public ConstrainedInt MaxMissedHeartbeats { get; set; } = new(
        3, 1, 20);

    /// <summary>
    ///     Debounce delay in milliseconds for session modification events.
    ///     Used globally for all extensions.
    /// </summary>
    public int DebounceDelayMs { get; set; } = 100;

    /// <summary>
    ///     Idle timeout configuration with default, constraints, and special value support.
    ///     Special value 0 means "never unload" (permanent resident).
    ///     Extensions can override within [Floor, Ceiling] range, or use 0 if allowed.
    /// </summary>
    public ConstrainedIntWithSpecial IdleTimeoutMinutes { get; set; } = new(
        30, 1, 1440,
        0);

    /// <summary>
    ///     Loads configuration from environment variables and command line arguments.
    ///     Command line arguments take precedence over environment variables.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>ExtensionConfig instance.</returns>
    public static ExtensionConfig LoadFromArgs(string[] args)
    {
        var config = new ExtensionConfig();
        config.LoadFromEnvironment();
        config.LoadFromCommandLine(args);
        return config;
    }

    /// <summary>
    ///     Loads configuration from environment variables.
    /// </summary>
    private void LoadFromEnvironment()
    {
        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_ENABLED"), out var enabled))
            Enabled = enabled;

        var configPath = Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_CONFIG");
        if (!string.IsNullOrEmpty(configPath))
            ConfigPath = configPath;

        var tempDir = Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_TEMP_DIR");
        if (!string.IsNullOrEmpty(tempDir))
            TempDirectory = tempDir;

        var transportMode = Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_TRANSPORT_MODE");
        if (!string.IsNullOrEmpty(transportMode))
            DefaultTransportMode = transportMode;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_HEALTH_INTERVAL"),
                out var healthInterval))
            HealthCheckIntervalSeconds = healthInterval;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_MAX_RESTARTS"), out var maxRestarts))
            MaxRestartAttempts = maxRestarts;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_RESTART_COOLDOWN"),
                out var restartCooldown))
            RestartCooldownSeconds = restartCooldown;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_GRACEFUL_SHUTDOWN_TIMEOUT"),
                out var gracefulShutdownTimeout))
            GracefulShutdownTimeoutSeconds = gracefulShutdownTimeout;

        if (long.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_MAX_SNAPSHOT_SIZE"),
                out var maxSnapshotSize))
            MaxSnapshotSizeBytes = maxSnapshotSize;

        if (long.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_MIN_FREE_DISK_SPACE"),
                out var minFreeDiskSpace))
            MinFreeDiskSpaceBytes = minFreeDiskSpace;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_FRAME_INTERVAL_DEFAULT"),
                out var frameIntervalDefault))
            FrameIntervalMs.Default = frameIntervalDefault;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_FRAME_INTERVAL_FLOOR"),
                out var frameIntervalFloor))
            FrameIntervalMs.Floor = frameIntervalFloor;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_FRAME_INTERVAL_CEILING"),
                out var frameIntervalCeiling))
            FrameIntervalMs.Ceiling = frameIntervalCeiling;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_SNAPSHOT_TTL_DEFAULT"),
                out var snapshotTtlDefault))
            SnapshotTtlSeconds.Default = snapshotTtlDefault;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_SNAPSHOT_TTL_FLOOR"),
                out var snapshotTtlFloor))
            SnapshotTtlSeconds.Floor = snapshotTtlFloor;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_SNAPSHOT_TTL_CEILING"),
                out var snapshotTtlCeiling))
            SnapshotTtlSeconds.Ceiling = snapshotTtlCeiling;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_MAX_MISSED_HEARTBEATS_DEFAULT"),
                out var maxMissedHeartbeatsDefault))
            MaxMissedHeartbeats.Default = maxMissedHeartbeatsDefault;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_MAX_MISSED_HEARTBEATS_FLOOR"),
                out var maxMissedHeartbeatsFloor))
            MaxMissedHeartbeats.Floor = maxMissedHeartbeatsFloor;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_MAX_MISSED_HEARTBEATS_CEILING"),
                out var maxMissedHeartbeatsCeiling))
            MaxMissedHeartbeats.Ceiling = maxMissedHeartbeatsCeiling;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_DEBOUNCE_DELAY"),
                out var debounceDelay))
            DebounceDelayMs = debounceDelay;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_IDLE_TIMEOUT_DEFAULT"),
                out var idleTimeoutDefault))
            IdleTimeoutMinutes.Default = idleTimeoutDefault;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_IDLE_TIMEOUT_FLOOR"),
                out var idleTimeoutFloor))
            IdleTimeoutMinutes.Floor = idleTimeoutFloor;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_IDLE_TIMEOUT_CEILING"),
                out var idleTimeoutCeiling))
            IdleTimeoutMinutes.Ceiling = idleTimeoutCeiling;

        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_EXTENSION_IDLE_TIMEOUT_SPECIAL_ALLOWED"),
                out var idleTimeoutSpecialAllowed))
            IdleTimeoutMinutes.SpecialAllowed = idleTimeoutSpecialAllowed;
    }

    /// <summary>
    ///     Loads configuration from command line arguments.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    private void LoadFromCommandLine(string[] args)
    {
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            ProcessBooleanArgs(arg);
            ProcessSystemIntArgs(arg, args, ref i);
            ProcessSystemLongArgs(arg, args, ref i);
            ProcessSystemStringArgs(arg, args, ref i);
            ProcessConstrainedIntArgs(arg, args, ref i);
        }
    }

    /// <summary>
    ///     Processes boolean command line arguments.
    /// </summary>
    private void ProcessBooleanArgs(string arg)
    {
        if (arg.Equals("--extension-enabled", StringComparison.OrdinalIgnoreCase))
            Enabled = true;
        else if (arg.Equals("--extension-disabled", StringComparison.OrdinalIgnoreCase))
            Enabled = false;
        else if (arg.Equals("--extension-idle-timeout-special-allowed", StringComparison.OrdinalIgnoreCase))
            IdleTimeoutMinutes.SpecialAllowed = true;
        else if (arg.Equals("--extension-idle-timeout-special-disallowed", StringComparison.OrdinalIgnoreCase))
            IdleTimeoutMinutes.SpecialAllowed = false;
    }

    /// <summary>
    ///     Processes system-level integer command line arguments.
    /// </summary>
    private void ProcessSystemIntArgs(string arg, string[] args, ref int index)
    {
        if (TryParseIntArg(arg, "--extension-health-interval", args, ref index, out var healthInterval))
            HealthCheckIntervalSeconds = healthInterval;
        else if (TryParseIntArg(arg, "--extension-max-restarts", args, ref index, out var maxRestarts))
            MaxRestartAttempts = maxRestarts;
        else if (TryParseIntArg(arg, "--extension-restart-cooldown", args, ref index, out var restartCooldown))
            RestartCooldownSeconds = restartCooldown;
        else if (TryParseIntArg(arg, "--extension-graceful-shutdown-timeout", args, ref index,
                     out var gracefulShutdownTimeout))
            GracefulShutdownTimeoutSeconds = gracefulShutdownTimeout;
        else if (TryParseIntArg(arg, "--extension-max-conversion-failures", args, ref index,
                     out var maxConversionFailures))
            MaxConversionFailures = maxConversionFailures;
        else if (TryParseIntArg(arg, "--extension-failure-backoff", args, ref index, out var failureBackoff))
            FailureBackoffSeconds = failureBackoff;
        else if (TryParseIntArg(arg, "--extension-conversion-cache-ttl", args, ref index, out var conversionCacheTtl))
            ConversionCacheTtlSeconds = conversionCacheTtl;
        else if (TryParseIntArg(arg, "--extension-max-cache-size", args, ref index, out var maxCacheSize))
            MaxConversionCacheSize = maxCacheSize;
    }

    /// <summary>
    ///     Processes system-level long integer command line arguments.
    /// </summary>
    private void ProcessSystemLongArgs(string arg, string[] args, ref int index)
    {
        if (TryParseLongArg(arg, "--extension-max-snapshot-size", args, ref index, out var maxSnapshotSize))
            MaxSnapshotSizeBytes = maxSnapshotSize;
        else if (TryParseLongArg(arg, "--extension-min-free-disk-space", args, ref index, out var minFreeDiskSpace))
            MinFreeDiskSpaceBytes = minFreeDiskSpace;
    }

    /// <summary>
    ///     Processes system-level string command line arguments.
    /// </summary>
    private void ProcessSystemStringArgs(string arg, string[] args, ref int index)
    {
        if (TryParseStringArg(arg, "--extension-config", args, ref index, out var configPath))
            ConfigPath = configPath;
        else if (TryParseStringArg(arg, "--extension-temp-dir", args, ref index, out var tempDir))
            TempDirectory = tempDir;
        else if (TryParseStringArg(arg, "--extension-transport-mode", args, ref index, out var transportMode))
            DefaultTransportMode = transportMode;
    }

    /// <summary>
    ///     Processes constrained integer command line arguments.
    /// </summary>
    private void ProcessConstrainedIntArgs(string arg, string[] args, ref int index)
    {
        if (TryParseIntArg(arg, "--extension-frame-interval", args, ref index, out var frameInterval))
            FrameIntervalMs.Default = frameInterval;
        else if (TryParseIntArg(arg, "--extension-frame-interval-default", args, ref index,
                     out var frameIntervalDefault))
            FrameIntervalMs.Default = frameIntervalDefault;
        else if (TryParseIntArg(arg, "--extension-frame-interval-floor", args, ref index, out var frameIntervalFloor))
            FrameIntervalMs.Floor = frameIntervalFloor;
        else if (TryParseIntArg(arg, "--extension-frame-interval-ceiling", args, ref index,
                     out var frameIntervalCeiling))
            FrameIntervalMs.Ceiling = frameIntervalCeiling;
        else if (TryParseIntArg(arg, "--extension-snapshot-ttl", args, ref index, out var snapshotTtl))
            SnapshotTtlSeconds.Default = snapshotTtl;
        else if (TryParseIntArg(arg, "--extension-snapshot-ttl-default", args, ref index, out var snapshotTtlDefault))
            SnapshotTtlSeconds.Default = snapshotTtlDefault;
        else if (TryParseIntArg(arg, "--extension-snapshot-ttl-floor", args, ref index, out var snapshotTtlFloor))
            SnapshotTtlSeconds.Floor = snapshotTtlFloor;
        else if (TryParseIntArg(arg, "--extension-snapshot-ttl-ceiling", args, ref index, out var snapshotTtlCeiling))
            SnapshotTtlSeconds.Ceiling = snapshotTtlCeiling;
        else if (TryParseIntArg(arg, "--extension-max-missed-heartbeats", args, ref index, out var maxMissedHeartbeats))
            MaxMissedHeartbeats.Default = maxMissedHeartbeats;
        else if (TryParseIntArg(arg, "--extension-max-missed-heartbeats-default", args, ref index,
                     out var maxMissedHeartbeatsDefault))
            MaxMissedHeartbeats.Default = maxMissedHeartbeatsDefault;
        else if (TryParseIntArg(arg, "--extension-max-missed-heartbeats-floor", args, ref index,
                     out var maxMissedHeartbeatsFloor))
            MaxMissedHeartbeats.Floor = maxMissedHeartbeatsFloor;
        else if (TryParseIntArg(arg, "--extension-max-missed-heartbeats-ceiling", args, ref index,
                     out var maxMissedHeartbeatsCeiling))
            MaxMissedHeartbeats.Ceiling = maxMissedHeartbeatsCeiling;
        else if (TryParseIntArg(arg, "--extension-debounce-delay", args, ref index, out var debounceDelay))
            DebounceDelayMs = debounceDelay;
        else if (TryParseIntArg(arg, "--extension-idle-timeout", args, ref index, out var idleTimeout))
            IdleTimeoutMinutes.Default = idleTimeout;
        else if (TryParseIntArg(arg, "--extension-idle-timeout-default", args, ref index, out var idleTimeoutDefault))
            IdleTimeoutMinutes.Default = idleTimeoutDefault;
        else if (TryParseIntArg(arg, "--extension-idle-timeout-floor", args, ref index, out var idleTimeoutFloor))
            IdleTimeoutMinutes.Floor = idleTimeoutFloor;
        else if (TryParseIntArg(arg, "--extension-idle-timeout-ceiling", args, ref index, out var idleTimeoutCeiling))
            IdleTimeoutMinutes.Ceiling = idleTimeoutCeiling;
    }

    /// <summary>
    ///     Tries to parse an integer argument with space, colon, and equals separators.
    /// </summary>
    private static bool TryParseIntArg(string arg, string prefix, string[] args, ref int index, out int value)
    {
        value = 0;

        if (arg.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
            index + 1 < args.Length && int.TryParse(args[index + 1], out value))
        {
            index++;
            return true;
        }

        var colonPrefix = prefix + ":";
        var equalsPrefix = prefix + "=";

        if (arg.StartsWith(colonPrefix, StringComparison.OrdinalIgnoreCase))
            return int.TryParse(arg[colonPrefix.Length..], out value);
        if (arg.StartsWith(equalsPrefix, StringComparison.OrdinalIgnoreCase))
            return int.TryParse(arg[equalsPrefix.Length..], out value);

        return false;
    }

    /// <summary>
    ///     Tries to parse a long integer argument with space, colon, and equals separators.
    /// </summary>
    private static bool TryParseLongArg(string arg, string prefix, string[] args, ref int index, out long value)
    {
        value = 0;

        if (arg.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
            index + 1 < args.Length && long.TryParse(args[index + 1], out value))
        {
            index++;
            return true;
        }

        var colonPrefix = prefix + ":";
        var equalsPrefix = prefix + "=";

        if (arg.StartsWith(colonPrefix, StringComparison.OrdinalIgnoreCase))
            return long.TryParse(arg[colonPrefix.Length..], out value);
        if (arg.StartsWith(equalsPrefix, StringComparison.OrdinalIgnoreCase))
            return long.TryParse(arg[equalsPrefix.Length..], out value);

        return false;
    }

    /// <summary>
    ///     Tries to parse a string argument with space, colon, and equals separators.
    /// </summary>
    private static bool TryParseStringArg(string arg, string prefix, string[] args, ref int index, out string value)
    {
        value = string.Empty;

        if (arg.Equals(prefix, StringComparison.OrdinalIgnoreCase) && index + 1 < args.Length)
        {
            value = args[index + 1];
            index++;
            return true;
        }

        var colonPrefix = prefix + ":";
        var equalsPrefix = prefix + "=";

        if (arg.StartsWith(colonPrefix, StringComparison.OrdinalIgnoreCase))
        {
            value = arg[colonPrefix.Length..];
            return true;
        }

        if (arg.StartsWith(equalsPrefix, StringComparison.OrdinalIgnoreCase))
        {
            value = arg[equalsPrefix.Length..];
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Validates the extension configuration.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when configuration is invalid.</exception>
    public void Validate()
    {
        if (!Enabled)
            return;

        if (HealthCheckIntervalSeconds < 1)
            throw new InvalidOperationException("HealthCheckIntervalSeconds must be at least 1");

        if (HealthCheckIntervalSeconds > 3600)
            throw new InvalidOperationException("HealthCheckIntervalSeconds cannot exceed 3600 (1 hour)");

        if (MaxRestartAttempts < 0)
            throw new InvalidOperationException("MaxRestartAttempts cannot be negative");

        if (MaxRestartAttempts > 100)
            throw new InvalidOperationException("MaxRestartAttempts cannot exceed 100");

        if (RestartCooldownSeconds < 0)
            throw new InvalidOperationException("RestartCooldownSeconds cannot be negative");

        if (ErrorRecoveryCooldownSeconds < 0)
            throw new InvalidOperationException("ErrorRecoveryCooldownSeconds cannot be negative");

        if (ErrorRecoveryCooldownSeconds > 3600)
            throw new InvalidOperationException("ErrorRecoveryCooldownSeconds cannot exceed 3600 (1 hour)");

        if (GracefulShutdownTimeoutSeconds < 1)
            throw new InvalidOperationException("GracefulShutdownTimeoutSeconds must be at least 1");

        if (StdinWriteTimeoutMs < 1000)
            throw new InvalidOperationException("StdinWriteTimeoutMs must be at least 1000");

        if (StdinWriteTimeoutMs > 60000)
            throw new InvalidOperationException("StdinWriteTimeoutMs cannot exceed 60000 (1 minute)");

        if (MaxConsecutiveSendFailures < 0)
            throw new InvalidOperationException("MaxConsecutiveSendFailures cannot be negative");

        if (RetryLoopTimeoutSeconds < 5)
            throw new InvalidOperationException("RetryLoopTimeoutSeconds must be at least 5");

        if (RetryLoopTimeoutSeconds > 300)
            throw new InvalidOperationException("RetryLoopTimeoutSeconds cannot exceed 300 (5 minutes)");

        if (ConversionCacheTtlSeconds < 1)
            throw new InvalidOperationException("ConversionCacheTtlSeconds must be at least 1");

        if (MaxConversionCacheSize < 1)
            throw new InvalidOperationException("MaxConversionCacheSize must be at least 1");

        if (MaxConversionCacheSize > 10000)
            throw new InvalidOperationException("MaxConversionCacheSize cannot exceed 10000");

        if (MaxConversionFailures < 1)
            throw new InvalidOperationException("MaxConversionFailures must be at least 1");

        if (FailureBackoffSeconds < 0)
            throw new InvalidOperationException("FailureBackoffSeconds cannot be negative");

        if (FailureBackoffSeconds > 3600)
            throw new InvalidOperationException("FailureBackoffSeconds cannot exceed 3600 (1 hour)");

        if (MaxSnapshotSizeBytes < 1024 * 1024)
            throw new InvalidOperationException("MaxSnapshotSizeBytes must be at least 1 MB (1048576 bytes)");

        if (MaxSnapshotSizeBytes > 1024L * 1024 * 1024)
            throw new InvalidOperationException("MaxSnapshotSizeBytes cannot exceed 1 GB (1073741824 bytes)");

        if (MinFreeDiskSpaceBytes < 0)
            throw new InvalidOperationException("MinFreeDiskSpaceBytes cannot be negative");

        if (MinFreeDiskSpaceBytes > 10L * 1024 * 1024 * 1024)
            throw new InvalidOperationException("MinFreeDiskSpaceBytes cannot exceed 10 GB");

        if (string.IsNullOrEmpty(TempDirectory))
            throw new InvalidOperationException("TempDirectory cannot be empty");

        ValidateTempDirectory(TempDirectory);

        var validTransportModes = new[] { "mmap", "stdin", "file" };
        if (!validTransportModes.Contains(DefaultTransportMode.ToLowerInvariant()))
            throw new InvalidOperationException(
                $"DefaultTransportMode must be one of: {string.Join(", ", validTransportModes)}");

        ValidateConstrainedInt(FrameIntervalMs, "FrameIntervalMs", 1, 60000);
        ValidateConstrainedInt(SnapshotTtlSeconds, "SnapshotTtlSeconds", 1, 600);
        ValidateConstrainedInt(MaxMissedHeartbeats, "MaxMissedHeartbeats", 1, 100);
        ValidateConstrainedInt(IdleTimeoutMinutes, "IdleTimeoutMinutes", 1, 10080); // 7 days

        if (DebounceDelayMs < 0)
            throw new InvalidOperationException("DebounceDelayMs cannot be negative");

        if (DebounceDelayMs > 10000)
            throw new InvalidOperationException("DebounceDelayMs cannot exceed 10000 (10 seconds)");
    }

    /// <summary>
    ///     Validates the extension configuration with session config dependency check.
    /// </summary>
    /// <param name="sessionConfig">The session configuration to validate against.</param>
    /// <exception cref="InvalidOperationException">Thrown when configuration is invalid.</exception>
    public void Validate(SessionConfig sessionConfig)
    {
        Validate();

        if (!Enabled)
            return;

        if (!sessionConfig.Enabled)
            throw new InvalidOperationException(
                "Extension requires Session to be enabled. " +
                "Use --session-enabled together with --extension-enabled");
    }

    /// <summary>
    ///     Validates a constrained integer setting.
    /// </summary>
    private static void ValidateConstrainedInt(ConstrainedInt setting, string name, int minFloor, int maxCeiling)
    {
        if (setting.Floor < minFloor)
            throw new InvalidOperationException($"{name}.Floor must be at least {minFloor}");

        if (setting.Ceiling > maxCeiling)
            throw new InvalidOperationException($"{name}.Ceiling cannot exceed {maxCeiling}");

        if (setting.Floor > setting.Ceiling)
            throw new InvalidOperationException($"{name}.Floor cannot be greater than {name}.Ceiling");

        if (setting.Default < setting.Floor || setting.Default > setting.Ceiling)
            throw new InvalidOperationException($"{name}.Default must be between Floor and Ceiling");
    }

    /// <summary>
    ///     Validates that the temp directory is safe and writable.
    /// </summary>
    private static void ValidateTempDirectory(string tempDir)
    {
        var fullPath = Path.GetFullPath(tempDir);
        var normalizedPath = fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            .ToLowerInvariant();

        var forbiddenPaths = GetForbiddenPaths();
        if (forbiddenPaths.Any(forbidden =>
                normalizedPath.Equals(forbidden, StringComparison.OrdinalIgnoreCase) ||
                normalizedPath.StartsWith(forbidden + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException(
                $"TempDirectory cannot be set to system directory: {tempDir}. " +
                "Please use a user-writable directory like the system temp folder.");

        try
        {
            Directory.CreateDirectory(fullPath);

            var testFile = Path.Combine(fullPath, $".write_test_{Guid.NewGuid():N}");
            File.WriteAllText(testFile, "test");
            File.Delete(testFile);
        }
        catch (UnauthorizedAccessException)
        {
            throw new InvalidOperationException(
                $"TempDirectory is not writable: {tempDir}. " +
                "Please ensure the application has write permissions.");
        }
        catch (IOException ex)
        {
            throw new InvalidOperationException(
                $"TempDirectory is not accessible: {tempDir}. Error: {ex.Message}");
        }
    }

    /// <summary>
    ///     Gets a list of forbidden system paths that should not be used as temp directories.
    /// </summary>
    private static List<string> GetForbiddenPaths()
    {
        var forbidden = new List<string>();

        if (OperatingSystem.IsWindows())
        {
            var winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            var system = Environment.GetFolderPath(Environment.SpecialFolder.System);
            var systemX86 = Environment.GetFolderPath(Environment.SpecialFolder.SystemX86);

            if (!string.IsNullOrEmpty(winDir)) forbidden.Add(winDir.ToLowerInvariant());
            if (!string.IsNullOrEmpty(programFiles)) forbidden.Add(programFiles.ToLowerInvariant());
            if (!string.IsNullOrEmpty(programFilesX86)) forbidden.Add(programFilesX86.ToLowerInvariant());
            if (!string.IsNullOrEmpty(system)) forbidden.Add(system.ToLowerInvariant());
            if (!string.IsNullOrEmpty(systemX86)) forbidden.Add(systemX86.ToLowerInvariant());
        }
        else
        {
            forbidden.AddRange(["/bin", "/sbin", "/usr/bin", "/usr/sbin", "/usr/lib", "/lib", "/etc", "/boot"]);
        }

        return forbidden;
    }
}
