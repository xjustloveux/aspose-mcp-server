namespace AsposeMcpServer.Core.Session;

/// <summary>
///     Session information for API responses
/// </summary>
public class SessionInfo
{
    /// <summary>
    ///     Unique session identifier
    /// </summary>
    public string SessionId { get; set; } = "";

    /// <summary>
    ///     Document type (word, excel, powerpoint, pdf)
    /// </summary>
    public string DocumentType { get; set; } = "";

    /// <summary>
    ///     Original file path
    /// </summary>
    public string Path { get; set; } = "";

    /// <summary>
    ///     Access mode (readonly, readwrite)
    /// </summary>
    public string Mode { get; set; } = "";

    /// <summary>
    ///     Whether the document has unsaved changes
    /// </summary>
    public bool IsDirty { get; set; }

    /// <summary>
    ///     When the session was opened
    /// </summary>
    public DateTime OpenedAt { get; set; }

    /// <summary>
    ///     Last access time
    /// </summary>
    public DateTime LastAccessedAt { get; set; }

    /// <summary>
    ///     Estimated memory usage in MB
    /// </summary>
    public double EstimatedMemoryMb { get; set; }
}
