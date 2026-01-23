using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Common;

/// <summary>
///     Generic success result with a message.
/// </summary>
public record SuccessResult
{
    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
