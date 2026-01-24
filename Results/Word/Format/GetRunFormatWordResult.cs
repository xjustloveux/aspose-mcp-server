using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Format;

/// <summary>
///     Result for getting run format information for a specific run.
/// </summary>
public sealed record GetRunFormatWordResult : RunFormatInfoBase
{
    /// <summary>
    ///     Zero-based paragraph index.
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    public required int ParagraphIndex { get; init; }

    /// <summary>
    ///     Format type (explicit or inherited).
    /// </summary>
    [JsonPropertyName("formatType")]
    public required string FormatType { get; init; }

    /// <summary>
    ///     Font name.
    /// </summary>
    [JsonPropertyName("fontName")]
    public required string FontName { get; init; }

    /// <summary>
    ///     Indicates whether the color is auto (empty or black).
    /// </summary>
    [JsonPropertyName("isAutoColor")]
    public required bool IsAutoColor { get; init; }
}
