namespace AsposeMcpServer.Core.Session;

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
    ///     Auto-save interval in minutes for dirty sessions.
    ///     When set to a value greater than 0, dirty sessions will be periodically saved to temp files.
    ///     This helps prevent data loss in case of unexpected termination (e.g., kill -9).
    ///     Default is 0 (disabled). Set via ASPOSE_SESSION_AUTO_SAVE_INTERVAL or --session-auto-save:N
    /// </summary>
    public int AutoSaveIntervalMinutes { get; set; }

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

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_SESSION_AUTO_SAVE_INTERVAL"),
                out var autoSaveInterval))
            AutoSaveIntervalMinutes = autoSaveInterval;
    }

    /// <summary>
    ///     Loads configuration from command line arguments (overrides environment variables)
    /// </summary>
    /// <param name="args">Command line arguments</param>
    private void LoadFromCommandLine(string[] args)
    {
        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            ProcessBooleanArgs(arg);
            ProcessIntArgs(arg, args, ref i);
            ProcessStringArgs(arg, args, ref i);
            ProcessEnumArgs(arg, args, ref i);
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
    /// <param name="arg">The current argument string.</param>
    /// <param name="args">The full arguments array.</param>
    /// <param name="index">The current index in the arguments array.</param>
    private void ProcessIntArgs(string arg, string[] args, ref int index)
    {
        if (TryParseIntArg(arg, "--session-max", args, ref index, out var maxSessions))
            MaxSessions = maxSessions;
        else if (TryParseIntArg(arg, "--session-timeout", args, ref index, out var timeout))
            IdleTimeoutMinutes = timeout;
        else if (TryParseIntArg(arg, "--session-max-file-size", args, ref index, out var maxFileSize))
            MaxFileSizeMb = maxFileSize;
        else if (TryParseIntArg(arg, "--session-temp-retention-hours", args, ref index, out var retention))
            TempRetentionHours = retention;
        else if (TryParseIntArg(arg, "--session-auto-save", args, ref index, out var autoSave))
            AutoSaveIntervalMinutes = autoSave;
    }

    /// <summary>
    ///     Processes string command line arguments.
    /// </summary>
    /// <param name="arg">The current argument string.</param>
    /// <param name="args">The full arguments array.</param>
    /// <param name="index">The current index in the arguments array.</param>
    private void ProcessStringArgs(string arg, string[] args, ref int index)
    {
        if (TryParseStringArg(arg, "--session-temp-dir", args, ref index, out var tempDir))
            TempDirectory = tempDir;
    }

    /// <summary>
    ///     Processes enum command line arguments.
    /// </summary>
    /// <param name="arg">The current argument string.</param>
    /// <param name="args">The full arguments array.</param>
    /// <param name="index">The current index in the arguments array.</param>
    private void ProcessEnumArgs(string arg, string[] args, ref int index)
    {
        if (TryParseEnumArg<DisconnectBehavior>(arg, "--session-on-disconnect", args, ref index, out var behavior))
            OnDisconnect = behavior;
        else if (TryParseEnumArg<SessionIsolationMode>(arg, "--session-isolation", args, ref index, out var isolation))
            IsolationMode = isolation;
    }

    /// <summary>
    ///     Tries to parse an integer argument with space, colon, and equals separators.
    /// </summary>
    /// <param name="arg">The argument string to parse.</param>
    /// <param name="prefix">The argument prefix to match.</param>
    /// <param name="args">The full arguments array for space-separated value lookup.</param>
    /// <param name="index">The current index in the arguments array.</param>
    /// <param name="value">The parsed integer value.</param>
    /// <returns>True if parsing succeeded; otherwise, false.</returns>
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
    ///     Tries to parse a string argument with space, colon, and equals separators.
    /// </summary>
    /// <param name="arg">The argument string to parse.</param>
    /// <param name="prefix">The argument prefix to match.</param>
    /// <param name="args">The full arguments array for space-separated value lookup.</param>
    /// <param name="index">The current index in the arguments array.</param>
    /// <param name="value">The parsed string value.</param>
    /// <returns>True if parsing succeeded; otherwise, false.</returns>
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
    ///     Tries to parse an enum argument with space, colon, and equals separators.
    /// </summary>
    /// <typeparam name="T">The enum type to parse.</typeparam>
    /// <param name="arg">The argument string to parse.</param>
    /// <param name="prefix">The argument prefix to match.</param>
    /// <param name="args">The full arguments array for space-separated value lookup.</param>
    /// <param name="index">The current index in the arguments array.</param>
    /// <param name="value">The parsed enum value.</param>
    /// <returns>True if parsing succeeded; otherwise, false.</returns>
    private static bool TryParseEnumArg<T>(string arg, string prefix, string[] args, ref int index, out T value)
        where T : struct, Enum
    {
        value = default;

        if (arg.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
            index + 1 < args.Length && Enum.TryParse(args[index + 1], true, out value))
        {
            index++;
            return true;
        }

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

        if (AutoSaveIntervalMinutes < 0)
            throw new InvalidOperationException("AutoSaveIntervalMinutes cannot be negative");

        if (string.IsNullOrEmpty(TempDirectory))
            throw new InvalidOperationException("TempDirectory cannot be empty");
    }
}
