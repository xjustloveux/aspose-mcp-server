using System.Text.Json;

namespace AsposeMcpServer.Core.Tasks;

/// <summary>
///     Information about an async task.
/// </summary>
public sealed class TaskInfo : IDisposable
{
    private bool _disposed;

    /// <summary>
    ///     Gets the unique task identifier.
    /// </summary>
    public required string TaskId { get; init; }

    /// <summary>
    ///     Gets or sets the current task status.
    /// </summary>
    public TaskStatus Status { get; set; } = TaskStatus.Working;

    /// <summary>
    ///     Gets or sets the human-readable status message.
    /// </summary>
    public string? StatusMessage { get; set; }

    /// <summary>
    ///     Gets when the task was created.
    /// </summary>
    public DateTime CreatedAt { get; init; } = DateTime.UtcNow;

    /// <summary>
    ///     Gets or sets when the task was last updated.
    /// </summary>
    public DateTime LastUpdatedAt { get; set; } = DateTime.UtcNow;

    /// <summary>
    ///     Gets or sets the time-to-live in milliseconds.
    /// </summary>
    public int Ttl { get; set; }

    /// <summary>
    ///     Gets or sets the suggested poll interval in milliseconds.
    /// </summary>
    public int PollInterval { get; set; } = 5000;

    /// <summary>
    ///     Gets the tool that was called.
    /// </summary>
    public required string ToolName { get; init; }

    /// <summary>
    ///     Gets the arguments passed to the tool.
    /// </summary>
    public required JsonElement Arguments { get; init; }

    /// <summary>
    ///     Gets or sets the task result (when completed).
    /// </summary>
    public string? Result { get; set; }

    /// <summary>
    ///     Gets or sets the error message (when failed).
    /// </summary>
    public string? ErrorMessage { get; set; }

    /// <summary>
    ///     Gets the owner identity (for session isolation).
    /// </summary>
    public string? OwnerId { get; init; }

    /// <summary>
    ///     Gets the cancellation token source for this task.
    /// </summary>
    internal CancellationTokenSource CancellationTokenSource { get; } = new();

    /// <summary>
    ///     Gets whether the task is in a terminal state.
    /// </summary>
    public bool IsTerminal => Status is TaskStatus.Completed or TaskStatus.Failed or TaskStatus.Cancelled;

    /// <summary>
    ///     Disposes the task resources.
    /// </summary>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        CancellationTokenSource.Dispose();
    }
}
