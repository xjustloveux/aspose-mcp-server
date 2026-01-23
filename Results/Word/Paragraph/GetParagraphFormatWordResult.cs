using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Result for getting paragraph format from Word documents.
/// </summary>
public sealed record GetParagraphFormatWordResult
{
    /// <summary>
    ///     Zero-based paragraph index.
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    public required int ParagraphIndex { get; init; }

    /// <summary>
    ///     Text content of the paragraph.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Length of the text content.
    /// </summary>
    [JsonPropertyName("textLength")]
    public required int TextLength { get; init; }

    /// <summary>
    ///     Number of runs in the paragraph.
    /// </summary>
    [JsonPropertyName("runCount")]
    public required int RunCount { get; init; }

    /// <summary>
    ///     Paragraph format information.
    /// </summary>
    [JsonPropertyName("paragraphFormat")]
    public required ParagraphFormatInfo ParagraphFormat { get; init; }

    /// <summary>
    ///     List format information (if paragraph is a list item).
    /// </summary>
    [JsonPropertyName("listFormat")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public ParagraphListFormatInfo? ListFormat { get; init; }

    /// <summary>
    ///     Border information (if any borders are set).
    /// </summary>
    [JsonPropertyName("borders")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyDictionary<string, BorderInfo>? Borders { get; init; }

    /// <summary>
    ///     Background color in hex format (#RRGGBB).
    /// </summary>
    [JsonPropertyName("backgroundColor")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? BackgroundColor { get; init; }

    /// <summary>
    ///     Tab stops (if any are defined).
    /// </summary>
    [JsonPropertyName("tabStops")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<ParagraphTabStopInfo>? TabStops { get; init; }

    /// <summary>
    ///     Font format of the first run.
    /// </summary>
    [JsonPropertyName("fontFormat")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public FontFormatInfo? FontFormat { get; init; }

    /// <summary>
    ///     Run details (when includeRunDetails is true and there are multiple runs).
    /// </summary>
    [JsonPropertyName("runs")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public RunDetailsInfo? Runs { get; init; }
}
