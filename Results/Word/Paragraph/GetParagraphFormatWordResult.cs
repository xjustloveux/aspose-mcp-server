using System.Text.Json.Serialization;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Result for getting paragraph format from Word documents.
/// </summary>
public sealed record GetParagraphFormatWordResult
{
    /// <summary>
    ///     Story-relative paragraph index (0-based within its story).
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    public required int ParagraphIndex { get; init; }

    /// <summary>
    ///     The story the paragraph belongs to (Body, Header, Footer, TextBox, Comment, Footnote, Endnote).
    /// </summary>
    [JsonPropertyName("storyType")]
    public string StoryType { get; init; } = StoryTypes.Body;

    /// <summary>
    ///     Section index for Body/Header/Footer stories.
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    public int SectionIndex { get; init; }

    /// <summary>
    ///     For Header/Footer stories: Primary, First, or Even; null otherwise.
    /// </summary>
    [JsonPropertyName("headerFooterType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? HeaderFooterType { get; init; }

    /// <summary>
    ///     Instance selector for multi-instance stories (e.g. Comment id); null when not applicable.
    /// </summary>
    [JsonPropertyName("containerIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ContainerIndex { get; init; }

    /// <summary>
    ///     Global document-order index across all stories (for cross-reference / disambiguation).
    /// </summary>
    [JsonPropertyName("documentOrderIndex")]
    public int DocumentOrderIndex { get; init; }

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
