using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.Content;

/// <summary>
///     Represents a single email header name-value pair.
/// </summary>
public record EmailHeaderInfo
{
    /// <summary>
    ///     The header name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     The header value.
    /// </summary>
    [JsonPropertyName("value")]
    public required string Value { get; init; }
}
