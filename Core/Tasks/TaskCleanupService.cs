namespace AsposeMcpServer.Core.Tasks;

/// <summary>
///     Background service that periodically cleans up expired tasks.
/// </summary>
public sealed class TaskCleanupService : BackgroundService
{
    private readonly TaskConfig _config;
    private readonly ILogger<TaskCleanupService>? _logger;
    private readonly TaskStore _store;

    /// <summary>
    ///     Creates a new task cleanup service.
    /// </summary>
    /// <param name="store">Task store.</param>
    /// <param name="config">Task configuration.</param>
    /// <param name="logger">Optional logger.</param>
    /// <exception cref="ArgumentNullException">Thrown when store or config is null.</exception>
    public TaskCleanupService(
        TaskStore store,
        TaskConfig config,
        ILogger<TaskCleanupService>? logger = null)
    {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _logger = logger;
    }

    /// <summary>
    ///     Executes the cleanup loop.
    /// </summary>
    /// <param name="stoppingToken">Cancellation token.</param>
    /// <returns>A task representing the background operation.</returns>
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger?.LogInformation("Task cleanup service started");

        while (!stoppingToken.IsCancellationRequested)
            try
            {
                await Task.Delay(_config.CleanupIntervalMs, stoppingToken);
                _store.CleanupExpiredTasks();
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error during task cleanup");
            }

        _logger?.LogInformation("Task cleanup service stopped");
    }
}
