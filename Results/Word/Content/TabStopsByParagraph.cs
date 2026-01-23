using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Content;

/// <summary>
///     Tab stops information for a specific paragraph.
/// </summary>
public sealed record TabStopsByParagraph
{
    /// <summary>
    ///     Zero-based section index.
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    public required int SectionIndex { get; init; }

    /// <summary>
    ///     Zero-based paragraph index within the section.
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    public required int ParagraphIndex { get; init; }

    /// <summary>
    ///     List of tab stops for this paragraph.
    /// </summary>
    [JsonPropertyName("tabStops")]
    public required IReadOnlyList<TabStopInfo> TabStops { get; init; }
}
