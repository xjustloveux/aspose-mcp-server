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
        {
            ProcessBooleanArgs(arg);
            ProcessIntArgs(arg);
            ProcessStringArgs(arg);
            ProcessEnumArgs(arg);
        }
    }

    /// <summary>
    ///     Processes boolean command line arguments.
    /// </summary>
    private void ProcessBooleanArgs(string arg)
    {
        if (arg.Equals("--session-enabled", StringComparison.OrdinalIgnoreCase))
            Enabled = true;
        else if (arg.Equals("--session-disabled", StringComparison.OrdinalIgnoreCase))
            Enabled = false;
    }

    /// <summary>
    ///     Processes integer command line arguments.
    /// </summary>
    private void ProcessIntArgs(string arg)
    {
        if (TryParseIntArg(arg, "--session-max", out var maxSessions))
            MaxSessions = maxSessions;
        else if (TryParseIntArg(arg, "--session-timeout", out var timeout))
            IdleTimeoutMinutes = timeout;
        else if (TryParseIntArg(arg, "--session-max-file-size", out var maxFileSize))
            MaxFileSizeMb = maxFileSize;
        else if (TryParseIntArg(arg, "--session-temp-retention-hours", out var retention))
            TempRetentionHours = retention;
    }

    /// <summary>
    ///     Processes string command line arguments.
    /// </summary>
    private void ProcessStringArgs(string arg)
    {
        if (TryParseStringArg(arg, "--session-temp-dir", out var tempDir))
            TempDirectory = tempDir;
    }

    /// <summary>
    ///     Processes enum command line arguments.
    /// </summary>
    private void ProcessEnumArgs(string arg)
    {
        if (TryParseEnumArg<DisconnectBehavior>(arg, "--session-on-disconnect", out var behavior))
            OnDisconnect = behavior;
        else if (TryParseEnumArg<SessionIsolationMode>(arg, "--session-isolation", out var isolation))
            IsolationMode = isolation;
    }

    /// <summary>
    ///     Tries to parse an integer argument with both : and = separators.
    /// </summary>
    /// <param name="arg">The argument string to parse.</param>
    /// <param name="prefix">The argument prefix to match.</param>
    /// <param name="value">The parsed integer value.</param>
    /// <returns>True if parsing succeeded; otherwise, false.</returns>
    private static bool TryParseIntArg(string arg, string prefix, out int value)
    {
        value = 0;
        var colonPrefix = prefix + ":";
        var equalsPrefix = prefix + "=";

        if (arg.StartsWith(colonPrefix, StringComparison.OrdinalIgnoreCase))
            return int.TryParse(arg[colonPrefix.Length..], out value);
        if (arg.StartsWith(equalsPrefix, StringComparison.OrdinalIgnoreCase))
            return int.TryParse(arg[equalsPrefix.Length..], out value);

        return false;
    }

    /// <summary>
    ///     Tries to parse a string argument with both : and = separators.
    /// </summary>
    /// <param name="arg">The argument string to parse.</param>
    /// <param name="prefix">The argument prefix to match.</param>
    /// <param name="value">The parsed string value.</param>
    /// <returns>True if parsing succeeded; otherwise, false.</returns>
    private static bool TryParseStringArg(string arg, string prefix, out string value)
    {
        value = string.Empty;
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
    ///     Tries to parse an enum argument with both : and = separators.
    /// </summary>
    /// <typeparam name="T">The enum type to parse.</typeparam>
    /// <param name="arg">The argument string to parse.</param>
    /// <param name="prefix">The argument prefix to match.</param>
    /// <param name="value">The parsed enum value.</param>
    /// <returns>True if parsing succeeded; otherwise, false.</returns>
    private static bool TryParseEnumArg<T>(string arg, string prefix, out T value) where T : struct, Enum
    {
        value = default;
        var colonPrefix = prefix + ":";
        var equalsPrefix = prefix + "=";

        if (arg.StartsWith(colonPrefix, StringComparison.OrdinalIgnoreCase))
            return Enum.TryParse(arg[colonPrefix.Length..], true, out value);
        if (arg.StartsWith(equalsPrefix, StringComparison.OrdinalIgnoreCase))
            return Enum.TryParse(arg[equalsPrefix.Length..], true, out value);

        return false;
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
