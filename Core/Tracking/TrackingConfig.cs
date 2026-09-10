namespace AsposeMcpServer.Core.Tracking;

/// <summary>
///     Tracking and monitoring configuration
/// </summary>
public class TrackingConfig
{
    /// <summary>
    ///     Routes the metrics path must not equal or contain.
    ///     <para>
    ///         Serving metrics anonymously is an opt-in for one endpoint. If that endpoint is
    ///         also a real route, or a parent of one, the exemption stops being about metrics:
    ///         a metrics path of <c>/</c> or <c>/mcp</c> would describe the tool endpoint too.
    ///     </para>
    /// </summary>
    private static readonly string[] ReservedRoutes = ["/mcp", "/health", "/ready"];

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
    ///     Whether the metrics endpoint is subject to the configured authentication.
    ///     Metrics expose request volume, success and failure counts, average latency, per-tool
    ///     usage and process memory, so the endpoint is authenticated like any other by default.
    ///     Set to <c>false</c> only when the endpoint is reachable exclusively from a trusted
    ///     scrape network, since liveness probes use <c>/health</c> and <c>/ready</c> instead.
    /// </summary>
    public bool MetricsRequireAuth { get; set; } = true;

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
            MetricsPath = "/" + MetricsPath;

        ValidateMetricsPathDoesNotCollide();
    }

    /// <summary>
    ///     Refuses a metrics path that would also describe another route.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when the path collides.</exception>
    private void ValidateMetricsPathDoesNotCollide()
    {
        if (string.IsNullOrEmpty(MetricsPath)) return;

        var normalized = MetricsPath.TrimEnd('/');
        if (normalized.Length == 0)
            throw new InvalidOperationException(
                "The metrics path cannot be the site root: every request would match it. "
                + "Use a dedicated path such as /metrics.");

        foreach (var reserved in ReservedRoutes)
        {
            var collides = normalized.Equals(reserved, StringComparison.OrdinalIgnoreCase)
                           || reserved.StartsWith(normalized + "/", StringComparison.OrdinalIgnoreCase)
                           || normalized.StartsWith(reserved + "/", StringComparison.OrdinalIgnoreCase);

            if (collides)
                throw new InvalidOperationException(
                    $"The metrics path '{MetricsPath}' collides with the reserved route "
                    + $"'{reserved}'. Anonymous metrics access would then cover that route as "
                    + "well. Use a dedicated path such as /metrics.");
        }
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

        if (bool.TryParse(Environment.GetEnvironmentVariable("ASPOSE_METRICS_REQUIRE_AUTH"), out var requireAuth))
            MetricsRequireAuth = requireAuth;
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

            if (arg.Equals("--log-enabled", StringComparison.OrdinalIgnoreCase))
            {
                LogEnabled = true;
            }
            else if (arg.Equals("--log-disabled", StringComparison.OrdinalIgnoreCase))
            {
                LogEnabled = false;
            }
            else if (TryGetStringValue(arg, "--log-targets", args, ref i, out var logTargets))
            {
                ParseLogTargets(logTargets);
            }
            else if (arg.Equals("--webhook-enabled", StringComparison.OrdinalIgnoreCase))
            {
                WebhookEnabled = true;
            }
            else if (arg.Equals("--webhook-disabled", StringComparison.OrdinalIgnoreCase))
            {
                WebhookEnabled = false;
            }
            else if (TryGetStringValue(arg, "--webhook-url", args, ref i, out var webhookUrl))
            {
                WebhookUrl = webhookUrl;
                if (!string.IsNullOrEmpty(WebhookUrl))
                    WebhookEnabled = true;
            }
            else if (TryGetIntValue(arg, "--webhook-timeout", args, ref i, out var timeout))
            {
                WebhookTimeoutSeconds = timeout;
            }
            else if (TryGetStringValue(arg, "--webhook-auth-header", args, ref i, out var authHeader))
            {
                WebhookAuthHeader = authHeader;
            }
            else if (arg.Equals("--metrics-enabled", StringComparison.OrdinalIgnoreCase))
            {
                MetricsEnabled = true;
            }
            else if (arg.Equals("--metrics-disabled", StringComparison.OrdinalIgnoreCase))
            {
                MetricsEnabled = false;
            }
            else if (TryGetStringValue(arg, "--metrics-path", args, ref i, out var metricsPath))
            {
                MetricsPath = metricsPath;
            }
            else if (arg.Equals("--metrics-allow-anonymous", StringComparison.OrdinalIgnoreCase))
            {
                MetricsRequireAuth = false;
            }
            else if (arg.Equals("--metrics-require-auth", StringComparison.OrdinalIgnoreCase))
            {
                MetricsRequireAuth = true;
            }
        }
    }

    /// <summary>
    ///     Tries to extract a string value from a command line argument with space, colon, or equals separator.
    /// </summary>
    /// <param name="arg">The current argument string.</param>
    /// <param name="prefix">The argument prefix to match.</param>
    /// <param name="args">The full arguments array.</param>
    /// <param name="index">The current index in the arguments array.</param>
    /// <param name="value">The extracted string value.</param>
    /// <returns>True if the argument was matched and a value extracted; otherwise, false.</returns>
    private static bool TryGetStringValue(string arg, string prefix, string[] args, ref int index, out string value)
    {
        value = string.Empty;

        if (arg.Equals(prefix, StringComparison.OrdinalIgnoreCase) && index + 1 < args.Length)
        {
            value = args[index + 1];
            index++;
            return true;
        }

        var colonPrefix = prefix + ":";
        if (arg.StartsWith(colonPrefix, StringComparison.OrdinalIgnoreCase))
        {
            value = arg[colonPrefix.Length..];
            return true;
        }

        var equalsPrefix = prefix + "=";
        if (arg.StartsWith(equalsPrefix, StringComparison.OrdinalIgnoreCase))
        {
            value = arg[equalsPrefix.Length..];
            return true;
        }

        return false;
    }

    /// <summary>
    ///     Tries to extract an integer value from a command line argument with space, colon, or equals separator.
    /// </summary>
    /// <param name="arg">The current argument string.</param>
    /// <param name="prefix">The argument prefix to match.</param>
    /// <param name="args">The full arguments array.</param>
    /// <param name="index">The current index in the arguments array.</param>
    /// <param name="value">The parsed integer value.</param>
    /// <returns>True if the argument was matched and a value parsed; otherwise, false.</returns>
    private static bool TryGetIntValue(string arg, string prefix, string[] args, ref int index, out int value)
    {
        value = 0;

        if (arg.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
            index + 1 < args.Length && int.TryParse(args[index + 1], out value))
        {
            index++;
            return true;
        }

        var colonPrefix = prefix + ":";
        if (arg.StartsWith(colonPrefix, StringComparison.OrdinalIgnoreCase))
            return int.TryParse(arg[colonPrefix.Length..], out value);

        var equalsPrefix = prefix + "=";
        if (arg.StartsWith(equalsPrefix, StringComparison.OrdinalIgnoreCase))
            return int.TryParse(arg[equalsPrefix.Length..], out value);

        return false;
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
