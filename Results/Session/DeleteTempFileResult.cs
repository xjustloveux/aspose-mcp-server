using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for delete temp file operation.
/// </summary>
public record DeleteTempFileResult
{
    /// <summary>
    ///     Whether the operation was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     Session ID that was deleted.
    /// </summary>
    [JsonPropertyName("sessionId")]
    public required string SessionId { get; init; }

    /// <summary>
    ///     Human-readable message.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
