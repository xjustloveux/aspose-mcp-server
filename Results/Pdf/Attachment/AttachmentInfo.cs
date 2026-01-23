using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Attachment;

/// <summary>
///     Information about a single attachment.
/// </summary>
public record AttachmentInfo
{
    /// <summary>
    ///     Zero-based index of the attachment.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     File name of the attachment.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Description of the attachment.
    /// </summary>
    [JsonPropertyName("description")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Description { get; init; }

    /// <summary>
    ///     MIME type of the attachment.
    /// </summary>
    [JsonPropertyName("mimeType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? MimeType { get; init; }

    /// <summary>
    ///     Size of the attachment in bytes.
    /// </summary>
    [JsonPropertyName("size")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? Size { get; init; }

    /// <summary>
    ///     Creation date.
    /// </summary>
    [JsonPropertyName("creationDate")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CreationDate { get; init; }

    /// <summary>
    ///     Modification date.
    /// </summary>
    [JsonPropertyName("modificationDate")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ModificationDate { get; init; }
}
