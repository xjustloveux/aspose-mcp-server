using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Session information DTO for result types.
/// </summary>
public record SessionInfoDto
{
    /// <summary>
    ///     Unique session identifier.
    /// </summary>
    [JsonPropertyName("sessionId")]
    public required string SessionId { get; init; }

    /// <summary>
    ///     Document type (word, excel, powerpoint, pdf).
    /// </summary>
    [JsonPropertyName("documentType")]
    public required string DocumentType { get; init; }

    /// <summary>
    ///     Original file path.
    /// </summary>
    [JsonPropertyName("path")]
    public required string Path { get; init; }

    /// <summary>
    ///     Access mode (readonly, readwrite).
    /// </summary>
    [JsonPropertyName("mode")]
    public required string Mode { get; init; }

    /// <summary>
    ///     Whether the document has unsaved changes.
    /// </summary>
    [JsonPropertyName("isDirty")]
    public bool IsDirty { get; init; }

    /// <summary>
    ///     When the session was opened.
    /// </summary>
    [JsonPropertyName("openedAt")]
    public DateTime OpenedAt { get; init; }

    /// <summary>
    ///     Last access time.
    /// </summary>
    [JsonPropertyName("lastAccessedAt")]
    public DateTime LastAccessedAt { get; init; }

    /// <summary>
    ///     Estimated memory usage in MB.
    /// </summary>
    [JsonPropertyName("estimatedMemoryMB")]
    public double EstimatedMemoryMb { get; init; }
}
