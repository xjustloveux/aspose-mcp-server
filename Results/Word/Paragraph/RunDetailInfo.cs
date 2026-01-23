using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Individual run detail information.
/// </summary>
public sealed record RunDetailInfo
{
    /// <summary>
    ///     Zero-based index within the paragraph.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Text content of the run.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

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
}
