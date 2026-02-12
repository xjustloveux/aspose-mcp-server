using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.Extension;

/// <summary>
///     Response message from extension during initialization handshake.
/// </summary>
public class ExtensionInitializeResponse
{
    /// <summary>
    ///     Message type, must be "initialize_response".
    /// </summary>
    [JsonPropertyName("type")]
    public string Type { get; set; } = "initialize_response";

    /// <summary>
    ///     Display name of the extension. Required.
    /// </summary>
    [JsonPropertyName("name")]
    public string Name { get; set; } = string.Empty;

    /// <summary>
    ///     Version of the extension (semver format recommended). Required.
    /// </summary>
    [JsonPropertyName("version")]
    public string Version { get; set; } = string.Empty;

    /// <summary>
    ///     Localized title for display purposes. Optional.
    /// </summary>
    [JsonPropertyName("title")]
    public string? Title { get; set; }

    /// <summary>
    ///     Description of what the extension does. Optional.
    /// </summary>
    [JsonPropertyName("description")]
    public string? Description { get; set; }

    /// <summary>
    ///     Author information. Optional.
    /// </summary>
    [JsonPropertyName("author")]
    public string? Author { get; set; }

    /// <summary>
    ///     Website URL for the extension. Optional.
    /// </summary>
    [JsonPropertyName("websiteUrl")]
    public string? WebsiteUrl { get; set; }
}
