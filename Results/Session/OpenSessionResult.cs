using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for open session operation.
/// </summary>
public record OpenSessionResult
{
    /// <summary>
    ///     Whether the operation was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     The session ID for subsequent operations.
    /// </summary>
    [JsonPropertyName("sessionId")]
    public required string SessionId { get; init; }

    /// <summary>
    ///     Human-readable message.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }

    /// <summary>
    ///     Session information.
    /// </summary>
    [JsonPropertyName("session")]
    public required SessionInfoDto Session { get; init; }
}
