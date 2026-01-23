namespace AsposeMcpServer.Core.Tracking;

/// <summary>
///     Tracking event data structure
/// </summary>
public class TrackingEvent
{
    /// <summary>
    ///     Event timestamp in UTC
    /// </summary>
    public DateTime Timestamp { get; set; } = DateTime.UtcNow;

    /// <summary>
    ///     Group identifier (from authentication)
    /// </summary>
    public string? GroupId { get; set; }

    /// <summary>
    ///     User identifier (from JWT authentication)
    /// </summary>
    public string? UserId { get; set; }

    /// <summary>
    ///     Tool name that was invoked
    /// </summary>
    public string? Tool { get; set; }

    /// <summary>
    ///     Operation that was performed
    /// </summary>
    public string? Operation { get; set; }

    /// <summary>
    ///     Duration of the operation in milliseconds
    /// </summary>
    public long DurationMs { get; set; }

    /// <summary>
    ///     Whether the operation was successful
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    ///     Error message if operation failed
    /// </summary>
    public string? Error { get; set; }

    /// <summary>
    ///     Current session memory usage in MB
    /// </summary>
    public double SessionMemoryMb { get; set; }

    /// <summary>
    ///     Session ID if operation used session mode
    /// </summary>
    public string? SessionId { get; set; }

    /// <summary>
    ///     Request ID for correlation
    /// </summary>
    public string? RequestId { get; set; }
}
