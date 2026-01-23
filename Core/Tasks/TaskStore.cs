using System.Collections.Concurrent;
using System.Text.Json;

namespace AsposeMcpServer.Core.Tasks;

/// <summary>
///     Thread-safe in-memory task storage.
/// </summary>
public sealed class TaskStore
{
    private readonly TaskConfig _config;
    private readonly object _createLock = new();
    private readonly ILogger<TaskStore>? _logger;
    private readonly ConcurrentDictionary<string, TaskInfo> _tasks = new();

    /// <summary>
    ///     Creates a new task store.
    /// </summary>
    /// <param name="config">Task configuration.</param>
    /// <param name="logger">Optional logger.</param>
    /// <exception cref="ArgumentNullException">Thrown when config is null.</exception>
    public TaskStore(TaskConfig config, ILogger<TaskStore>? logger = null)
    {
        _config = config ?? throw new ArgumentNullException(nameof(config));
        _logger = logger;
    }

    /// <summary>
    ///     Gets the current number of active (working) tasks.
    /// </summary>
    public int ActiveTaskCount => _tasks.Values.Count(t => t.Status == TaskStatus.Working);

    /// <summary>
    ///     Gets the total number of tasks in the store.
    /// </summary>
    public int TotalTaskCount => _tasks.Count;

    /// <summary>
    ///     Creates a new task.
    /// </summary>
    /// <param name="toolName">Name of the tool being called.</param>
    /// <param name="arguments">Tool arguments.</param>
    /// <param name="ttlMs">Optional TTL override in milliseconds.</param>
    /// <param name="ownerId">Optional owner identity for isolation.</param>
    /// <returns>The created task info.</returns>
    /// <exception cref="ArgumentNullException">Thrown when toolName is null or empty.</exception>
    /// <exception cref="InvalidOperationException">When max concurrent tasks is reached.</exception>
    public TaskInfo CreateTask(
        string toolName,
        JsonElement arguments,
        int? ttlMs = null,
        string? ownerId = null)
    {
        if (string.IsNullOrEmpty(toolName))
            throw new ArgumentNullException(nameof(toolName));

        lock (_createLock)
        {
            var activeTasks = _tasks.Values.Count(t =>
                t.Status == TaskStatus.Working &&
                (ownerId == null || t.OwnerId == ownerId));

            if (activeTasks >= _config.MaxConcurrentTasks)
                throw new InvalidOperationException(
                    $"Maximum concurrent tasks ({_config.MaxConcurrentTasks}) reached. " +
                    "Please wait for existing tasks to complete or cancel them.");

            var taskId = Guid.NewGuid().ToString("N");
            var effectiveTtl = Math.Min(ttlMs ?? _config.DefaultTtlMs, _config.MaxTtlMs);

            var task = new TaskInfo
            {
                TaskId = taskId,
                ToolName = toolName,
                Arguments = arguments.Clone(),
                Ttl = effectiveTtl,
                PollInterval = _config.DefaultPollIntervalMs,
                OwnerId = ownerId,
                StatusMessage = "Task created, waiting to start"
            };

            if (!_tasks.TryAdd(taskId, task))
            {
                task.Dispose();
                throw new InvalidOperationException("Failed to create task: ID collision");
            }

            _logger?.LogInformation("Task {TaskId} created for tool {ToolName}", taskId, toolName);
            return task;
        }
    }

    /// <summary>
    ///     Gets a task by ID.
    /// </summary>
    /// <param name="taskId">The task ID.</param>
    /// <param name="ownerId">Optional owner ID for isolation.</param>
    /// <returns>The task info, or null if not found or access denied.</returns>
    public TaskInfo? GetTask(string taskId, string? ownerId = null)
    {
        if (string.IsNullOrEmpty(taskId))
            return null;

        if (!_tasks.TryGetValue(taskId, out var task))
            return null;

        if (ownerId != null && task.OwnerId != ownerId)
            return null;

        return task;
    }

    /// <summary>
    ///     Lists all tasks, optionally filtered by owner.
    /// </summary>
    /// <param name="ownerId">Optional owner ID for filtering.</param>
    /// <returns>List of tasks ordered by creation time (newest first).</returns>
    public IReadOnlyList<TaskInfo> ListTasks(string? ownerId = null)
    {
        var tasks = _tasks.Values.AsEnumerable();

        if (ownerId != null)
            tasks = tasks.Where(t => t.OwnerId == ownerId);

        return tasks.OrderByDescending(t => t.CreatedAt).ToList();
    }

    /// <summary>
    ///     Updates task status atomically.
    /// </summary>
    /// <param name="taskId">The task ID.</param>
    /// <param name="status">New status.</param>
    /// <param name="statusMessage">Optional status message.</param>
    /// <param name="result">Optional result (for completed tasks).</param>
    /// <param name="errorMessage">Optional error message (for failed tasks).</param>
    /// <returns>True if updated, false if task not found.</returns>
    public bool UpdateTaskStatus(
        string taskId,
        TaskStatus status,
        string? statusMessage = null,
        string? result = null,
        string? errorMessage = null)
    {
        if (!_tasks.TryGetValue(taskId, out var task))
            return false;

        task.Status = status;
        task.LastUpdatedAt = DateTime.UtcNow;

        if (statusMessage != null)
            task.StatusMessage = statusMessage;
        if (result != null)
            task.Result = result;
        if (errorMessage != null)
            task.ErrorMessage = errorMessage;

        _logger?.LogInformation("Task {TaskId} status updated to {Status}", taskId, status);
        return true;
    }

    /// <summary>
    ///     Cancels a task.
    /// </summary>
    /// <param name="taskId">The task ID.</param>
    /// <param name="ownerId">Optional owner ID for isolation.</param>
    /// <returns>True if cancelled, false if not found or already terminal.</returns>
    public bool CancelTask(string taskId, string? ownerId = null)
    {
        var task = GetTask(taskId, ownerId);
        if (task == null)
            return false;

        if (task.IsTerminal)
            return false;

        try
        {
            task.CancellationTokenSource.Cancel();
        }
        catch (ObjectDisposedException)
        {
            return false;
        }

        UpdateTaskStatus(taskId, TaskStatus.Cancelled, "Task cancelled by user");
        return true;
    }

    /// <summary>
    ///     Removes expired tasks.
    /// </summary>
    /// <returns>Number of tasks removed.</returns>
    public int CleanupExpiredTasks()
    {
        var now = DateTime.UtcNow;
        var expiredIds = _tasks
            .Where(kvp =>
                kvp.Value.IsTerminal &&
                now - kvp.Value.LastUpdatedAt > TimeSpan.FromMilliseconds(kvp.Value.Ttl))
            .Select(kvp => kvp.Key)
            .ToList();

        var removed = 0;
        foreach (var id in expiredIds)
            if (_tasks.TryRemove(id, out var task))
            {
                task.Dispose();
                removed++;
            }

        if (removed > 0)
            _logger?.LogInformation("Cleaned up {Count} expired tasks", removed);

        return removed;
    }

    /// <summary>
    ///     Removes all tasks (for testing or shutdown).
    /// </summary>
    public void Clear()
    {
        foreach (var kvp in _tasks)
            if (_tasks.TryRemove(kvp.Key, out var task))
                task.Dispose();
    }
}
