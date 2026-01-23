using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Session;

/// <summary>
///     Result for close session operation.
/// </summary>
public record CloseSessionResult
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
}
