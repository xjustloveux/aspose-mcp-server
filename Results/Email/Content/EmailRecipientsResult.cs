using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.Content;

/// <summary>
///     Result containing email recipient information.
/// </summary>
public record EmailRecipientsResult
{
    /// <summary>
    ///     The sender email address.
    /// </summary>
    [JsonPropertyName("from")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? From { get; init; }

    /// <summary>
    ///     The primary recipient (To) email addresses.
    /// </summary>
    [JsonPropertyName("to")]
    public required IReadOnlyList<string> To { get; init; }

    /// <summary>
    ///     The carbon copy (CC) email addresses.
    /// </summary>
    [JsonPropertyName("cc")]
    public required IReadOnlyList<string> Cc { get; init; }

    /// <summary>
    ///     The blind carbon copy (BCC) email addresses.
    /// </summary>
    [JsonPropertyName("bcc")]
    public required IReadOnlyList<string> Bcc { get; init; }

    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
