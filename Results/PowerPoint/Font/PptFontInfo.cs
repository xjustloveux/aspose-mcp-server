using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Font;

/// <summary>
///     Information about a font used in a presentation.
/// </summary>
public record PptFontInfo
{
    /// <summary>
    ///     The font name.
    /// </summary>
    [JsonPropertyName("fontName")]
    public required string FontName { get; init; }

    /// <summary>
    ///     Whether the font is embedded in the presentation.
    /// </summary>
    [JsonPropertyName("isEmbedded")]
    public required bool IsEmbedded { get; init; }

    /// <summary>
    ///     Whether the font is a custom font.
    /// </summary>
    [JsonPropertyName("isCustom")]
    public required bool IsCustom { get; init; }
}
