namespace AsposeMcpServer.Core.Tasks;

/// <summary>
///     Configuration for async task execution.
/// </summary>
public class TaskConfig
{
    /// <summary>
    ///     Whether async task execution is enabled. Default: true.
    /// </summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    ///     Maximum number of concurrent tasks per session. Default: 5.
    /// </summary>
    public int MaxConcurrentTasks { get; set; } = 5;

    /// <summary>
    ///     Default time-to-live for task results in milliseconds. Default: 300000 (5 minutes).
    /// </summary>
    public int DefaultTtlMs { get; set; } = 300000;

    /// <summary>
    ///     Maximum allowed TTL in milliseconds. Default: 3600000 (1 hour).
    /// </summary>
    public int MaxTtlMs { get; set; } = 3600000;

    /// <summary>
    ///     Default poll interval suggestion in milliseconds. Default: 5000 (5 seconds).
    /// </summary>
    public int DefaultPollIntervalMs { get; set; } = 5000;

    /// <summary>
    ///     Interval for cleanup of expired tasks in milliseconds. Default: 60000 (1 minute).
    /// </summary>
    public int CleanupIntervalMs { get; set; } = 60000;

    /// <summary>
    ///     Loads configuration from environment variables and command line arguments.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>Configured TaskConfig instance.</returns>
    public static TaskConfig LoadFromArgs(string[] args)
    {
        var config = new TaskConfig();

        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_TASKS_ENABLED"), out var enabled))
            config.Enabled = enabled;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_TASKS_MAX_CONCURRENT"), out var maxConcurrent))
            config.MaxConcurrentTasks = maxConcurrent;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_TASKS_DEFAULT_TTL"), out var defaultTtl))
            config.DefaultTtlMs = defaultTtl;

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_TASKS_MAX_TTL"), out var maxTtl))
            config.MaxTtlMs = maxTtl;

        foreach (var arg in args)
            if (arg.Equals("--no-tasks", StringComparison.OrdinalIgnoreCase))
            {
                config.Enabled = false;
            }
            else if (arg.StartsWith("--tasks-max-concurrent:", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(arg[23..], out var val))
                    config.MaxConcurrentTasks = val;
            }
            else if (arg.StartsWith("--tasks-default-ttl:", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(arg[20..], out var val))
                    config.DefaultTtlMs = val;
            }
            else if (arg.StartsWith("--tasks-max-ttl:", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(arg[16..], out var val))
                    config.MaxTtlMs = val;
            }

        return config;
    }

    /// <summary>
    ///     Validates the configuration values.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when configuration is invalid.</exception>
    public void Validate()
    {
        if (MaxConcurrentTasks < 1)
            throw new InvalidOperationException("MaxConcurrentTasks must be at least 1");

        if (MaxConcurrentTasks > 100)
            throw new InvalidOperationException("MaxConcurrentTasks cannot exceed 100");

        if (DefaultTtlMs < 1000)
            throw new InvalidOperationException("DefaultTtlMs must be at least 1000 (1 second)");

        if (MaxTtlMs < DefaultTtlMs)
            throw new InvalidOperationException("MaxTtlMs must be greater than or equal to DefaultTtlMs");

        if (CleanupIntervalMs < 1000)
            throw new InvalidOperationException("CleanupIntervalMs must be at least 1000 (1 second)");
    }
}
