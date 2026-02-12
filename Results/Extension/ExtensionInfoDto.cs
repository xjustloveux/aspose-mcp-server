using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Extension;

/// <summary>
///     Extension information for API responses.
/// </summary>
public record ExtensionInfoDto
{
    /// <summary>
    ///     Extension ID.
    /// </summary>
    [JsonPropertyName("id")]
    public required string Id { get; init; }

    /// <summary>
    ///     Display name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Version.
    /// </summary>
    [JsonPropertyName("version")]
    public required string Version { get; init; }

    /// <summary>
    ///     Localized title for display.
    /// </summary>
    [JsonPropertyName("title")]
    public string? Title { get; init; }

    /// <summary>
    ///     Description.
    /// </summary>
    [JsonPropertyName("description")]
    public string? Description { get; init; }

    /// <summary>
    ///     Author information.
    /// </summary>
    [JsonPropertyName("author")]
    public string? Author { get; init; }

    /// <summary>
    ///     Website URL.
    /// </summary>
    [JsonPropertyName("websiteUrl")]
    public string? WebsiteUrl { get; init; }

    /// <summary>
    ///     Whether the extension is available.
    /// </summary>
    [JsonPropertyName("isAvailable")]
    public bool IsAvailable { get; init; }

    /// <summary>
    ///     Whether the extension is currently initializing.
    /// </summary>
    [JsonPropertyName("isInitializing")]
    public bool IsInitializing { get; init; }

    /// <summary>
    ///     Reason if unavailable.
    /// </summary>
    [JsonPropertyName("unavailableReason")]
    public string? UnavailableReason { get; init; }

    /// <summary>
    ///     Supported document types.
    /// </summary>
    [JsonPropertyName("supportedDocumentTypes")]
    public required IReadOnlyList<string> SupportedDocumentTypes { get; init; }

    /// <summary>
    ///     Supported input formats.
    /// </summary>
    [JsonPropertyName("inputFormats")]
    public required IReadOnlyList<string> InputFormats { get; init; }

    /// <summary>
    ///     Current state (if loaded).
    /// </summary>
    [JsonPropertyName("state")]
    public string? State { get; init; }
}
