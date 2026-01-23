using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Attachment;

/// <summary>
///     Result for getting attachments from PDF documents.
/// </summary>
public record GetAttachmentsResult
{
    /// <summary>
    ///     Number of attachments.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of attachment information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<AttachmentInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no attachments found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
