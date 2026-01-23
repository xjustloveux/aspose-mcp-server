using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for save session operation.
/// </summary>
public record SaveSessionResult
{
    /// <summary>
    ///     Whether the operation was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     Human-readable message.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }

    /// <summary>
    ///     Session information after save.
    /// </summary>
    [JsonPropertyName("session")]
    public required SessionInfoDto Session { get; init; }
}
