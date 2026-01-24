using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Format;

/// <summary>
///     Base class for run format information.
/// </summary>
public record RunFormatInfoBase
{
    /// <summary>
    ///     Zero-based run index within the paragraph.
    /// </summary>
    [JsonPropertyName("runIndex")]
    public required int RunIndex { get; init; }

    /// <summary>
    ///     Text content of the run.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     ASCII font name.
    /// </summary>
    [JsonPropertyName("fontNameAscii")]
    public required string FontNameAscii { get; init; }

    /// <summary>
    ///     Far East font name.
    /// </summary>
    [JsonPropertyName("fontNameFarEast")]
    public required string FontNameFarEast { get; init; }

    /// <summary>
    ///     Font size in points.
    /// </summary>
    [JsonPropertyName("fontSize")]
    public required double FontSize { get; init; }

    /// <summary>
    ///     Indicates bold formatting.
    /// </summary>
    [JsonPropertyName("bold")]
    public required bool Bold { get; init; }

    /// <summary>
    ///     Indicates italic formatting.
    /// </summary>
    [JsonPropertyName("italic")]
    public required bool Italic { get; init; }

    /// <summary>
    ///     Underline style (None, Single, Double, etc.).
    /// </summary>
    [JsonPropertyName("underline")]
    public required string Underline { get; init; }

    /// <summary>
    ///     Indicates strikethrough formatting.
    /// </summary>
    [JsonPropertyName("strikeThrough")]
    public required bool StrikeThrough { get; init; }

    /// <summary>
    ///     Indicates superscript formatting.
    /// </summary>
    [JsonPropertyName("superscript")]
    public required bool Superscript { get; init; }

    /// <summary>
    ///     Indicates subscript formatting.
    /// </summary>
    [JsonPropertyName("subscript")]
    public required bool Subscript { get; init; }

    /// <summary>
    ///     Font color in hex format (#RRGGBB).
    /// </summary>
    [JsonPropertyName("color")]
    public required string Color { get; init; }

    /// <summary>
    ///     Color name.
    /// </summary>
    [JsonPropertyName("colorName")]
    public required string ColorName { get; init; }
}
