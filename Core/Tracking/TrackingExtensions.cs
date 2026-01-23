namespace AsposeMcpServer.Core.Tracking;

/// <summary>
///     Extension methods for adding tracking middleware
/// </summary>
public static class TrackingExtensions
{
    /// <summary>
    ///     Adds tracking middleware to the application pipeline
    /// </summary>
    /// <param name="app">The application builder</param>
    /// <param name="config">Tracking configuration</param>
    /// <returns>The application builder for chaining</returns>
    public static IApplicationBuilder UseTracking(
        this IApplicationBuilder app,
        TrackingConfig config)
    {
        if (config is { LogEnabled: false, WebhookEnabled: false, MetricsEnabled: false })
            return app;

        return app.UseMiddleware<TrackingMiddleware>(config);
    }
}
