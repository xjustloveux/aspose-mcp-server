namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Configuration for Origin header validation middleware.
///     Used to prevent DNS rebinding attacks in HTTP transport modes.
/// </summary>
public class OriginValidationConfig
{
    /// <summary>
    ///     Gets or sets whether Origin validation is enabled.
    ///     Default: true for HTTP transports.
    /// </summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    ///     Gets or sets whether to allow requests from localhost origins.
    ///     Default: true (for development).
    /// </summary>
    public bool AllowLocalhost { get; set; } = true;

    /// <summary>
    ///     Gets or sets whether to allow requests without Origin header.
    ///     Examples: curl, Postman, server-to-server requests.
    ///     Default: true.
    /// </summary>
    public bool AllowMissingOrigin { get; set; } = true;

    /// <summary>
    ///     Gets or sets the explicit list of allowed origins.
    ///     Example: ["https://myapp.example.com", "https://admin.example.com"]
    /// </summary>
    public string[]? AllowedOrigins { get; set; }

    /// <summary>
    ///     Gets or sets the paths to exclude from Origin validation.
    ///     Default: ["/health", "/ready"]
    /// </summary>
    public string[] ExcludedPaths { get; set; } = ["/health", "/ready"];

    /// <summary>
    ///     Loads configuration from environment variables and command line arguments.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>Configured OriginValidationConfig instance.</returns>
    public static OriginValidationConfig LoadFromArgs(string[] args)
    {
        var config = new OriginValidationConfig();

        if (bool.TryParse(
                Environment.GetEnvironmentVariable("ASPOSE_ORIGIN_VALIDATION"),
                out var enabled))
            config.Enabled = enabled;

        if (bool.TryParse(
                Environment.GetEnvironmentVariable("ASPOSE_ALLOW_LOCALHOST"),
                out var allowLocalhost))
            config.AllowLocalhost = allowLocalhost;

        if (bool.TryParse(
                Environment.GetEnvironmentVariable("ASPOSE_ALLOW_MISSING_ORIGIN"),
                out var allowMissing))
            config.AllowMissingOrigin = allowMissing;

        var allowedOrigins = Environment.GetEnvironmentVariable("ASPOSE_ALLOWED_ORIGINS");
        if (!string.IsNullOrEmpty(allowedOrigins))
            config.AllowedOrigins = allowedOrigins.Split(',', StringSplitOptions.RemoveEmptyEntries);

        foreach (var arg in args)
            if (arg.Equals("--no-origin-validation", StringComparison.OrdinalIgnoreCase))
                config.Enabled = false;
            else if (arg.Equals("--no-localhost", StringComparison.OrdinalIgnoreCase))
                config.AllowLocalhost = false;
            else if (arg.Equals("--require-origin", StringComparison.OrdinalIgnoreCase))
                config.AllowMissingOrigin = false;
            else if (arg.StartsWith("--allowed-origins:", StringComparison.OrdinalIgnoreCase))
                config.AllowedOrigins = arg[18..].Split(',', StringSplitOptions.RemoveEmptyEntries);

        return config;
    }
}
