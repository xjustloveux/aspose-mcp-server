using System.Text.Json.Serialization;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Results.Word.Format;

/// <summary>
///     Result for getting run format information for a specific run.
/// </summary>
public sealed record GetRunFormatWordResult : RunFormatInfoBase
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
