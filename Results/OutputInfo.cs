using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results;

/// <summary>
///     Unified output information structure for all tool results.
/// </summary>
public record OutputInfo
{
    /// <summary>
    ///     Output file path (for file mode operations).
    /// </summary>
    [JsonPropertyName("path")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Path { get; init; }

    /// <summary>
    ///     Session ID (for session mode operations).
    /// </summary>
    [JsonPropertyName("sessionId")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SessionId { get; init; }

    /// <summary>
    ///     Indicates whether the operation was performed in session mode.
    /// </summary>
    [JsonPropertyName("isSession")]
    public bool IsSession { get; init; }
}
