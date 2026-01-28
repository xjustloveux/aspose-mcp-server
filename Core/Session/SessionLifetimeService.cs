namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Hosted service that handles application lifecycle events for session management.
///     Ensures sessions are properly saved/cleaned up when the application shuts down.
/// </summary>
public class SessionLifetimeService : IHostedService
{
    private readonly IHostApplicationLifetime _lifetime;
    private readonly ILogger<SessionLifetimeService> _logger;
    private readonly DocumentSessionManager _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the SessionLifetimeService.
    /// </summary>
    /// <param name="sessionManager">The document session manager.</param>
    /// <param name="lifetime">The application lifetime service.</param>
    /// <param name="logger">The logger instance.</param>
    public SessionLifetimeService(
        DocumentSessionManager sessionManager,
        IHostApplicationLifetime lifetime,
        ILogger<SessionLifetimeService> logger)
    {
        _sessionManager = sessionManager;
        _lifetime = lifetime;
        _logger = logger;
    }

    /// <summary>
    ///     Starts the service and registers the application stopping event handler.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>A completed task.</returns>
    public Task StartAsync(CancellationToken cancellationToken)
    {
        _lifetime.ApplicationStopping.Register(OnApplicationStopping);
        _logger.LogDebug("SessionLifetimeService started, registered ApplicationStopping handler");
        return Task.CompletedTask;
    }

    /// <summary>
    ///     Stops the service.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>A completed task.</returns>
    public Task StopAsync(CancellationToken cancellationToken)
    {
        return Task.CompletedTask;
    }

    /// <summary>
    ///     Handles the application stopping event by triggering session cleanup.
    /// </summary>
    private void OnApplicationStopping()
    {
        _logger.LogInformation("Application stopping, triggering session cleanup");
        try
        {
            _sessionManager.OnServerShutdown();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error during session cleanup on application shutdown");
        }
    }
}
