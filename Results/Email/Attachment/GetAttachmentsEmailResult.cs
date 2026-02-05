using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.Attachment;

/// <summary>
///     Result of listing email attachments.
/// </summary>
public record GetAttachmentsEmailResult
{
    /// <summary>
    ///     Total number of attachments in the email.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of attachment information.
    /// </summary>
    [JsonPropertyName("attachments")]
    public required IReadOnlyList<AttachmentEmailInfo> Attachments { get; init; }

    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
