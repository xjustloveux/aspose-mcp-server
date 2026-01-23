namespace AsposeMcpServer.Core.Security;

/// <summary>
///     Middleware that validates Origin header to prevent DNS rebinding attacks.
///     Should be placed before authentication middleware in the pipeline.
/// </summary>
public class OriginValidationMiddleware
{
    private readonly HashSet<string> _allowedOrigins;
    private readonly OriginValidationConfig _config;
    private readonly HashSet<string> _excludedPaths;
    private readonly ILogger<OriginValidationMiddleware>? _logger;
    private readonly RequestDelegate _next;

    /// <summary>
    ///     Initializes a new instance of the <see cref="OriginValidationMiddleware" /> class.
    /// </summary>
    /// <param name="next">The next middleware in the pipeline.</param>
    /// <param name="config">Origin validation configuration.</param>
    /// <param name="logger">Optional logger for diagnostic output.</param>
    /// <exception cref="ArgumentNullException">Thrown when next or config is null.</exception>
    public OriginValidationMiddleware(
        RequestDelegate next,
        OriginValidationConfig config,
        ILogger<OriginValidationMiddleware>? logger = null)
    {
        _next = next ?? throw new ArgumentNullException(nameof(next));
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _logger = logger;
        _allowedOrigins = config.AllowedOrigins?.ToHashSet(StringComparer.OrdinalIgnoreCase)
                          ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        _excludedPaths = config.ExcludedPaths.ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     Processes the HTTP request and validates the Origin header.
    /// </summary>
    /// <param name="context">The HTTP context for the current request.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    public async Task InvokeAsync(HttpContext context)
    {
        if (!_config.Enabled || IsExcludedPath(context.Request.Path))
        {
            await _next(context);
            return;
        }

        var origin = context.Request.Headers.Origin.ToString();

        if (string.IsNullOrEmpty(origin))
        {
            if (_config.AllowMissingOrigin)
            {
                await _next(context);
                return;
            }

            _logger?.LogWarning(
                "Request rejected: Missing Origin header from {RemoteIp}",
                context.Connection.RemoteIpAddress);

            context.Response.StatusCode = StatusCodes.Status403Forbidden;
            context.Response.ContentType = "text/plain";
            await context.Response.WriteAsync("Origin header required");
            return;
        }

        if (IsAllowedOrigin(origin))
        {
            await _next(context);
            return;
        }

        _logger?.LogWarning(
            "Request rejected: Origin '{Origin}' not allowed from {RemoteIp}",
            origin,
            context.Connection.RemoteIpAddress);

        context.Response.StatusCode = StatusCodes.Status403Forbidden;
        context.Response.ContentType = "text/plain";
        await context.Response.WriteAsync("Origin not allowed");
    }

    /// <summary>
    ///     Checks if the request path is excluded from Origin validation.
    /// </summary>
    /// <param name="path">The request path to check.</param>
    /// <returns>True if the path is excluded; otherwise, false.</returns>
    private bool IsExcludedPath(PathString path)
    {
        foreach (var excludedPath in _excludedPaths)
            if (path.StartsWithSegments(excludedPath, StringComparison.OrdinalIgnoreCase))
                return true;

        return false;
    }

    /// <summary>
    ///     Checks if the Origin is in the allowed list or is a localhost origin.
    /// </summary>
    /// <param name="origin">The Origin header value to validate.</param>
    /// <returns>True if the origin is allowed; otherwise, false.</returns>
    private bool IsAllowedOrigin(string origin)
    {
        if (_allowedOrigins.Contains(origin)) return true;

        if (_config.AllowLocalhost && IsLocalhostOrigin(origin)) return true;

        return false;
    }

    /// <summary>
    ///     Determines whether the specified origin is a localhost origin.
    /// </summary>
    /// <param name="origin">The Origin header value to check.</param>
    /// <returns>True if the origin is localhost; otherwise, false.</returns>
    private static bool IsLocalhostOrigin(string origin)
    {
        try
        {
            var uri = new Uri(origin);
            var host = uri.Host.ToLowerInvariant();

            return host is "localhost" or "127.0.0.1" or "[::1]"
                   || host.EndsWith(".localhost", StringComparison.OrdinalIgnoreCase);
        }
        catch (UriFormatException)
        {
            return false;
        }
    }
}
