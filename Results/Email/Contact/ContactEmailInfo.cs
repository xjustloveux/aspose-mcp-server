using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Email.Contact;

/// <summary>
///     Result containing email contact information.
/// </summary>
public record ContactEmailInfo
{
    /// <summary>
    ///     Display name of the contact.
    /// </summary>
    [JsonPropertyName("displayName")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DisplayName { get; init; }

    /// <summary>
    ///     Primary email address of the contact.
    /// </summary>
    [JsonPropertyName("email")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Email { get; init; }

    /// <summary>
    ///     Primary phone number of the contact.
    /// </summary>
    [JsonPropertyName("phone")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Phone { get; init; }

    /// <summary>
    ///     Company name of the contact.
    /// </summary>
    [JsonPropertyName("company")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Company { get; init; }

    /// <summary>
    ///     Job title of the contact.
    /// </summary>
    [JsonPropertyName("jobTitle")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? JobTitle { get; init; }

    /// <summary>
    ///     Whether the contact has a photo attached.
    /// </summary>
    [JsonPropertyName("hasPhoto")]
    public bool HasPhoto { get; init; }

    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
