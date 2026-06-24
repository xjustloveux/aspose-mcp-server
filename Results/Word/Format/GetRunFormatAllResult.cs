using System.Text.Json.Serialization;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Results.Word.Format;

/// <summary>
///     Result for getting run format information for all runs in a paragraph.
/// </summary>
public sealed record GetRunFormatAllResult
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
    ///     Total number of runs.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of run format information.
    /// </summary>
    [JsonPropertyName("runs")]
    public required IReadOnlyList<RunFormatInfo> Runs { get; init; }
}
