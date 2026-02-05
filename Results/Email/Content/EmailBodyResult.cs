using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.Content;

/// <summary>
///     Result containing email body content.
/// </summary>
public record EmailBodyResult
{
    /// <summary>
    ///     The plain text body of the email.
    /// </summary>
    [JsonPropertyName("body")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Body { get; init; }

    /// <summary>
    ///     The HTML body of the email.
    /// </summary>
    [JsonPropertyName("htmlBody")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? HtmlBody { get; init; }

    /// <summary>
    ///     Whether the email body is primarily HTML.
    /// </summary>
    [JsonPropertyName("isHtml")]
    public bool IsHtml { get; init; }

    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
