namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Session isolation mode for configurable group-based isolation.
///     The group identifier is determined by the GROUP_IDENTIFIER_CLAIM configuration.
/// </summary>
public enum SessionIsolationMode
{
    /// <summary>
    ///     No isolation - all users can access all sessions (backward compatible with Stdio mode)
    /// </summary>
    None,

    /// <summary>
    ///     Group-level isolation - users within the same group can access each other's sessions.
    ///     The group is determined by the configured GROUP_IDENTIFIER_CLAIM (e.g., tenant_id, team_id, sub).
    /// </summary>
    Group
}

/// <summary>
///     Configuration for document session management
/// </summary>
public class SessionConfig
{
    /// <summary>
    ///     Whether session management is enabled
    /// </summary>
    public bool Enabled { get; set; }

    /// <summary>
    ///     Behavior when client disconnects
    /// </summary>
    public DisconnectBehavior OnDisconnect { get; set; } = DisconnectBehavior.SaveToTemp;

    /// <summary>
    ///     Idle timeout in minutes (0 = no timeout)
    /// </summary>
    public int IdleTimeoutMinutes { get; set; } = 30;

    /// <summary>
    ///     Temporary directory for SaveToTemp behavior
    /// </summary>
    public string TempDirectory { get; set; } = Path.GetTempPath();

    /// <summary>
    ///     Maximum number of concurrent sessions
    /// </summary>
    public int MaxSessions { get; set; } = 10;

    /// <summary>
    ///     Maximum file size in MB for session mode
    /// </summary>
    public int MaxFileSizeMb { get; set; } = 100;

    /// <summary>
    ///     Temp file retention in hours. Files older than this will be cleaned up.
    ///     Default is 24 hours. Set via ASPOSE_SESSION_TEMP_RETENTION_HOURS or --session-temp-retention-hours:N
    /// </summary>
    public int TempRetentionHours { get; set; } = 24;

    /// <summary>
    ///     Session isolation mode for group-based environments.
    ///     Default is Group (users within the same group can access each other's sessions).
    ///     The group is determined by GROUP_IDENTIFIER_CLAIM configuration.
    ///     Set via ASPOSE_SESSION_ISOLATION or --session-isolation:mode
    /// </summary>
    public SessionIsolationMode IsolationMode { get; set; } = SessionIsolationMode.Group;

    /// <summary>
    ///     Loads configuration from environment variables and command line arguments.
    ///     Command line arguments take precedence over environment variables.
    /// </summary>
    /// <param name="args">Command line arguments</param>
    /// <returns>SessionConfig instance</returns>
    public static SessionConfig LoadFromArgs(string[] args)
    {
        var config = new SessionConfig();
        config.LoadFromEnvironment();
        config.LoadFromCommandLine(args);
        return config;
    }

    /// <summary>
    ///     Loads configuration from environment variables
    /// </summary>
    private void LoadFromEnvironment()
    {
        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_SESSION_ENABLED"), out var enabled))
            Enabled = enabled;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_SESSION_MAX"), out var maxSessions))
            MaxSessions = maxSessions;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_SESSION_TIMEOUT"), out var timeout))
            IdleTimeoutMinutes = timeout;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_SESSION_MAX_FILE_SIZE_MB"), out var maxFileSizeMb))
            MaxFileSizeMb = maxFileSizeMb;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_SESSION_TEMP_RETENTION_HOURS"),
                out var tempRetentionHours))
            TempRetentionHours = tempRetentionHours;

        var tempDir = Environment.GetEnvironmentVariable("ASPOSE_SESSION_TEMP_DIR");
        if (!string.IsNullOrEmpty(tempDir))
            TempDirectory = tempDir;

        var onDisconnect = Environment.GetEnvironmentVariable("ASPOSE_SESSION_ON_DISCONNECT");
        if (!string.IsNullOrEmpty(onDisconnect) &&
            Enum.TryParse<DisconnectBehavior>(onDisconnect, true, out var behavior))
            OnDisconnect = behavior;

        var isolation = Environment.GetEnvironmentVariable("ASPOSE_SESSION_ISOLATION");
        if (!string.IsNullOrEmpty(isolation) &&
            Enum.TryParse<SessionIsolationMode>(isolation, true, out var isolationMode))
            IsolationMode = isolationMode;
    }

    /// <summary>
    ///     Loads configuration from command line arguments (overrides environment variables)
    /// </summary>
    /// <param name="args">Command line arguments</param>
    private void LoadFromCommandLine(string[] args)
    {
        foreach (var arg in args)
            if (arg.Equals("--session-enabled", StringComparison.OrdinalIgnoreCase))
                Enabled = true;
            else if (arg.Equals("--session-disabled", StringComparison.OrdinalIgnoreCase))
                Enabled = false;
            else if (arg.StartsWith("--session-max:", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--session-max:".Length..], out var max1))
                MaxSessions = max1;
            else if (arg.StartsWith("--session-max=", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--session-max=".Length..], out var max2))
                MaxSessions = max2;
            else if (arg.StartsWith("--session-timeout:", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--session-timeout:".Length..], out var timeout1))
                IdleTimeoutMinutes = timeout1;
            else if (arg.StartsWith("--session-timeout=", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--session-timeout=".Length..], out var timeout2))
                IdleTimeoutMinutes = timeout2;
            else if (arg.StartsWith("--session-max-file-size:", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--session-max-file-size:".Length..], out var size1))
                MaxFileSizeMb = size1;
            else if (arg.StartsWith("--session-max-file-size=", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--session-max-file-size=".Length..], out var size2))
                MaxFileSizeMb = size2;
            else if (arg.StartsWith("--session-temp-dir:", StringComparison.OrdinalIgnoreCase))
                TempDirectory = arg["--session-temp-dir:".Length..];
            else if (arg.StartsWith("--session-temp-dir=", StringComparison.OrdinalIgnoreCase))
                TempDirectory = arg["--session-temp-dir=".Length..];
            else if (arg.StartsWith("--session-temp-retention-hours:", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--session-temp-retention-hours:".Length..], out var retention1))
                TempRetentionHours = retention1;
            else if (arg.StartsWith("--session-temp-retention-hours=", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--session-temp-retention-hours=".Length..], out var retention2))
                TempRetentionHours = retention2;
            else if (arg.StartsWith("--session-on-disconnect:", StringComparison.OrdinalIgnoreCase) &&
                     Enum.TryParse<DisconnectBehavior>(arg["--session-on-disconnect:".Length..], true,
                         out var behavior1))
                OnDisconnect = behavior1;
            else if (arg.StartsWith("--session-on-disconnect=", StringComparison.OrdinalIgnoreCase) &&
                     Enum.TryParse<DisconnectBehavior>(arg["--session-on-disconnect=".Length..], true,
                         out var behavior2))
                OnDisconnect = behavior2;
            else if (arg.StartsWith("--session-isolation:", StringComparison.OrdinalIgnoreCase) &&
                     Enum.TryParse<SessionIsolationMode>(arg["--session-isolation:".Length..], true,
                         out var isolation1))
                IsolationMode = isolation1;
            else if (arg.StartsWith("--session-isolation=", StringComparison.OrdinalIgnoreCase) &&
                     Enum.TryParse<SessionIsolationMode>(arg["--session-isolation=".Length..], true,
                         out var isolation2))
                IsolationMode = isolation2;
    }

    /// <summary>
    ///     Validates the session configuration
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when configuration is invalid</exception>
    public void Validate()
    {
        if (!Enabled)
            return;

        if (MaxSessions < 1)
            throw new InvalidOperationException("MaxSessions must be at least 1");

        if (IdleTimeoutMinutes < 0)
            throw new InvalidOperationException("IdleTimeoutMinutes cannot be negative");

        if (MaxFileSizeMb < 1)
            throw new InvalidOperationException("MaxFileSizeMb must be at least 1");

        if (TempRetentionHours < 1)
            throw new InvalidOperationException("TempRetentionHours must be at least 1");

        if (string.IsNullOrEmpty(TempDirectory))
            throw new InvalidOperationException("TempDirectory cannot be empty");
    }
}
