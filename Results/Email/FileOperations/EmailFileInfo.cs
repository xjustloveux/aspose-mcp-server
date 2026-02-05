using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.FileOperations;

/// <summary>
///     Result containing email file metadata and basic information.
/// </summary>
public record EmailFileInfo
{
    /// <summary>
    ///     The email subject line.
    /// </summary>
    [JsonPropertyName("subject")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Subject { get; init; }

    /// <summary>
    ///     The sender email address.
    /// </summary>
    [JsonPropertyName("from")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? From { get; init; }

    /// <summary>
    ///     The primary recipient email addresses.
    /// </summary>
    [JsonPropertyName("to")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? To { get; init; }

    /// <summary>
    ///     The email date.
    /// </summary>
    [JsonPropertyName("date")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Date { get; init; }

    /// <summary>
    ///     The detected email format (EML, MSG, etc.).
    /// </summary>
    [JsonPropertyName("format")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Format { get; init; }

    /// <summary>
    ///     Whether the email has any attachments.
    /// </summary>
    [JsonPropertyName("hasAttachments")]
    public bool HasAttachments { get; init; }

    /// <summary>
    ///     The number of attachments in the email.
    /// </summary>
    [JsonPropertyName("attachmentCount")]
    public int AttachmentCount { get; init; }

    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
