using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Font format information.
/// </summary>
public sealed record FontFormatInfo
{
    /// <summary>
    ///     Font size in points.
    /// </summary>
    [JsonPropertyName("fontSize")]
    public required double FontSize { get; init; }

    /// <summary>
    ///     Font name (when ASCII and Far East fonts are the same).
    /// </summary>
    [JsonPropertyName("font")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Font { get; init; }

    /// <summary>
    ///     ASCII font name (when different from Far East).
    /// </summary>
    [JsonPropertyName("fontAscii")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontAscii { get; init; }

    /// <summary>
    ///     Far East font name (when different from ASCII).
    /// </summary>
    [JsonPropertyName("fontFarEast")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? FontFarEast { get; init; }

    /// <summary>
    ///     Indicates bold formatting.
    /// </summary>
    [JsonPropertyName("bold")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool Bold { get; init; }

    /// <summary>
    ///     Indicates italic formatting.
    /// </summary>
    [JsonPropertyName("italic")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool Italic { get; init; }

    /// <summary>
    ///     Underline style.
    /// </summary>
    [JsonPropertyName("underline")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Underline { get; init; }

    /// <summary>
    ///     Indicates strikethrough formatting.
    /// </summary>
    [JsonPropertyName("strikethrough")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool Strikethrough { get; init; }

    /// <summary>
    ///     Indicates superscript formatting.
    /// </summary>
    [JsonPropertyName("superscript")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool Superscript { get; init; }

    /// <summary>
    ///     Indicates subscript formatting.
    /// </summary>
    [JsonPropertyName("subscript")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool Subscript { get; init; }

    /// <summary>
    ///     Font color in hex format (#RRGGBB).
    /// </summary>
    [JsonPropertyName("color")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Color { get; init; }

    /// <summary>
    ///     Highlight color name.
    /// </summary>
    [JsonPropertyName("highlightColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? HighlightColor { get; init; }
}
