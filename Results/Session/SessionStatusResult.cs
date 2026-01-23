using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for session status operation.
/// </summary>
public record SessionStatusResult
{
    /// <summary>
    ///     Whether the operation was successful.
    /// </summary>
    [JsonPropertyName("success")]
    public bool Success { get; init; }

    /// <summary>
    ///     Session information.
    /// </summary>
    [JsonPropertyName("session")]
    public required SessionInfoDto Session { get; init; }
}
