using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Attachment;

/// <summary>
///     Information about a single attachment.
/// </summary>
public record AttachmentInfo
{
    /// <summary>
    ///     One-based index of the attachment within the PDF embedded-file collection.
    ///     <para>
    ///         For display and ordering only. <c>pdf_attachment delete</c> selects by
    ///         <c>attachmentName</c> and has no index parameter, so following this field as if it
    ///         were the delete selector produced a request the tool rejects (R2-C03). Where two
    ///         attachments share a name, deletion by name cannot distinguish them.
    ///     </para>
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
