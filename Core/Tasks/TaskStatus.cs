namespace AsposeMcpServer.Core.Tasks;

/// <summary>
///     Task execution status as defined by MCP specification (2025-11-25).
/// </summary>
public enum TaskStatus
{
    /// <summary>Task is currently being processed.</summary>
    Working,

    /// <summary>Task is awaiting input from the requestor.</summary>
    InputRequired,

    /// <summary>Task completed successfully.</summary>
    Completed,

    /// <summary>Task execution failed.</summary>
    Failed,

    /// <summary>Task was cancelled by the requestor.</summary>
    Cancelled
}
