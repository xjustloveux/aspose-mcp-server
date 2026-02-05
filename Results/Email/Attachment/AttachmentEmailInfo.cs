using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.Attachment;

/// <summary>
///     Information about a single email attachment.
/// </summary>
public record AttachmentEmailInfo
{
    /// <summary>
    ///     Zero-based index of the attachment in the email.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     File name of the attachment.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     MIME content type of the attachment.
    /// </summary>
    [JsonPropertyName("contentType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ContentType { get; init; }

    /// <summary>
    ///     Size of the attachment in bytes.
    /// </summary>
    [JsonPropertyName("size")]
    public required long Size { get; init; }
}
