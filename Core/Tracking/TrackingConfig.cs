namespace AsposeMcpServer.Core.Tracking;

/// <summary>
///     Tracking and monitoring configuration
/// </summary>
public class TrackingConfig
{
    /// <summary>
    ///     Enable structured logging
    /// </summary>
    public bool LogEnabled { get; set; } = true;

    /// <summary>
    ///     Log output targets
    /// </summary>
    public LogTarget[] LogTargets { get; set; } = [LogTarget.Console];

    /// <summary>
    ///     Enable webhook notifications
    /// </summary>
    public bool WebhookEnabled { get; set; }

    /// <summary>
    ///     Webhook URL for event notifications
    /// </summary>
    public string? WebhookUrl { get; set; }

    /// <summary>
    ///     Authorization header value for webhook calls
    /// </summary>
    public string? WebhookAuthHeader { get; set; }

    /// <summary>
    ///     Timeout in seconds for webhook calls
    /// </summary>
    public int WebhookTimeoutSeconds { get; set; } = 5;

    /// <summary>
    ///     Enable Prometheus metrics endpoint
    /// </summary>
    public bool MetricsEnabled { get; set; }

    /// <summary>
    ///     Path for metrics endpoint
    /// </summary>
    public string MetricsPath { get; set; } = "/metrics";

    /// <summary>
    ///     Loads configuration from environment variables and command line arguments.
    ///     Command line arguments take precedence over environment variables.
    /// </summary>
    /// <param name="args">Command line arguments</param>
    /// <returns>TrackingConfig instance</returns>
    public static TrackingConfig LoadFromArgs(string[] args)
    {
        var config = new TrackingConfig();
        config.LoadFromEnvironment();
        config.LoadFromCommandLine(args);
        return config;
    }

    /// <summary>
    ///     Validates the configuration values
    /// </summary>
    public void Validate()
    {
        if (WebhookTimeoutSeconds is < 1 or > 300)
        {
            Console.Error.WriteLine($"[WARN] Invalid webhook timeout {WebhookTimeoutSeconds}, using default 5");
            WebhookTimeoutSeconds = 5;
        }

        if (!string.IsNullOrEmpty(MetricsPath) && !MetricsPath.StartsWith('/'))
            MetricsPath = "/" + MetricsPath; // NOSONAR S1075 - URL path prefix, not file system path
    }

    /// <summary>
    ///     Loads configuration from environment variables
    /// </summary>
    private void LoadFromEnvironment()
    {
        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_LOG_ENABLED"), out var logEnabled))
            LogEnabled = logEnabled;

        var logTargets = Environment.GetEnvironmentVariable("ASPOSE_LOG_TARGETS");
        if (!string.IsNullOrEmpty(logTargets))
            ParseLogTargets(logTargets);

        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_WEBHOOK_ENABLED"), out var webhookEnabled))
            WebhookEnabled = webhookEnabled;

        WebhookUrl = Environment.GetEnvironmentVariable("ASPOSE_WEBHOOK_URL");
        if (!string.IsNullOrEmpty(WebhookUrl) && !WebhookEnabled)
            WebhookEnabled = true; // Auto-enable if URL is provided

        WebhookAuthHeader = Environment.GetEnvironmentVariable("ASPOSE_WEBHOOK_AUTH_HEADER");

        if (int.TryParse(Environment.GetEnvironmentVariable("ASPOSE_WEBHOOK_TIMEOUT"), out var webhookTimeout))
            WebhookTimeoutSeconds = webhookTimeout;

        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_METRICS_ENABLED"), out var metricsEnabled))
            MetricsEnabled = metricsEnabled;

        var metricsPath = Environment.GetEnvironmentVariable("ASPOSE_METRICS_PATH");
        if (!string.IsNullOrEmpty(metricsPath))
            MetricsPath = metricsPath;
    }

    /// <summary>
    ///     Loads configuration from command line arguments (overrides environment variables)
    /// </summary>
    /// <param name="args">Command line arguments</param>
    private void
        LoadFromCommandLine(string[] args)
    {
        foreach (var arg in args)
            if (arg.Equals("--log-enabled", StringComparison.OrdinalIgnoreCase))
            {
                LogEnabled = true;
            }
            else if (arg.Equals("--log-disabled", StringComparison.OrdinalIgnoreCase))
            {
                LogEnabled = false;
            }
            else if (arg.StartsWith("--log-targets:", StringComparison.OrdinalIgnoreCase))
            {
                ParseLogTargets(arg["--log-targets:".Length..]);
            }
            else if (arg.StartsWith("--log-targets=", StringComparison.OrdinalIgnoreCase))
            {
                ParseLogTargets(arg["--log-targets=".Length..]);
            }
            else if (arg.Equals("--webhook-enabled", StringComparison.OrdinalIgnoreCase))
            {
                WebhookEnabled = true;
            }
            else if (arg.Equals("--webhook-disabled", StringComparison.OrdinalIgnoreCase))
            {
                WebhookEnabled = false;
            }
            else if (arg.StartsWith("--webhook-url:", StringComparison.OrdinalIgnoreCase))
            {
                WebhookUrl = arg["--webhook-url:".Length..];
                if (!string.IsNullOrEmpty(WebhookUrl))
                    WebhookEnabled = true;
            }
            else if (arg.StartsWith("--webhook-url=", StringComparison.OrdinalIgnoreCase))
            {
                WebhookUrl = arg["--webhook-url=".Length..];
                if (!string.IsNullOrEmpty(WebhookUrl))
                    WebhookEnabled = true;
            }
            else if (arg.StartsWith("--webhook-timeout:", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--webhook-timeout:".Length..], out var timeout1))
            {
                WebhookTimeoutSeconds = timeout1;
            }
            else if (arg.StartsWith("--webhook-timeout=", StringComparison.OrdinalIgnoreCase) &&
                     int.TryParse(arg["--webhook-timeout=".Length..], out var timeout2))
            {
                WebhookTimeoutSeconds = timeout2;
            }
            else if (arg.StartsWith("--webhook-auth-header:", StringComparison.OrdinalIgnoreCase))
            {
                WebhookAuthHeader = arg["--webhook-auth-header:".Length..];
            }
            else if (arg.StartsWith("--webhook-auth-header=", StringComparison.OrdinalIgnoreCase))
            {
                WebhookAuthHeader = arg["--webhook-auth-header=".Length..];
            }
            else if (arg.Equals("--metrics-enabled", StringComparison.OrdinalIgnoreCase))
            {
                MetricsEnabled = true;
            }
            else if (arg.Equals("--metrics-disabled", StringComparison.OrdinalIgnoreCase))
            {
                MetricsEnabled = false;
            }
            else if (arg.StartsWith("--metrics-path:", StringComparison.OrdinalIgnoreCase))
            {
                MetricsPath = arg["--metrics-path:".Length..];
            }
            else if (arg.StartsWith("--metrics-path=", StringComparison.OrdinalIgnoreCase))
            {
                MetricsPath = arg["--metrics-path=".Length..];
            }
    }

    /// <summary>
    ///     Parses log targets string in format "Console,EventLog"
    /// </summary>
    /// <param name="targetsString">Comma-separated log target names</param>
    private void ParseLogTargets(string targetsString)
    {
        List<LogTarget> targets = [];
        foreach (var target in targetsString.Split(',', StringSplitOptions.RemoveEmptyEntries))
            if (Enum.TryParse<LogTarget>(target.Trim(), true, out var parsedTarget))
                targets.Add(parsedTarget);

        if (targets.Count > 0)
            LogTargets = targets.ToArray();
    }
}
