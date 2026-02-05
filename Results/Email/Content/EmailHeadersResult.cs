using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.Content;

/// <summary>
///     Result containing email headers information.
/// </summary>
public record EmailHeadersResult
{
    /// <summary>
    ///     The list of email headers.
    /// </summary>
    [JsonPropertyName("headers")]
    public required IReadOnlyList<EmailHeaderInfo> Headers { get; init; }

    /// <summary>
    ///     The total number of headers.
    /// </summary>
    [JsonPropertyName("count")]
    public int Count { get; init; }

    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
