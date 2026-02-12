using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Binding information for API responses.
/// </summary>
public record BindingInfoDto
{
    /// <summary>
    ///     Session ID.
    /// </summary>
    [JsonPropertyName("sessionId")]
    public required string SessionId { get; init; }

    /// <summary>
    ///     Extension ID.
    /// </summary>
    [JsonPropertyName("extensionId")]
    public required string ExtensionId { get; init; }

    /// <summary>
    ///     Output format.
    /// </summary>
    [JsonPropertyName("outputFormat")]
    public required string OutputFormat { get; init; }

    /// <summary>
    ///     When the binding was created.
    /// </summary>
    [JsonPropertyName("createdAt")]
    public DateTime CreatedAt { get; init; }

    /// <summary>
    ///     When the last snapshot was sent.
    /// </summary>
    [JsonPropertyName("lastSentAt")]
    public DateTime? LastSentAt { get; init; }
}
