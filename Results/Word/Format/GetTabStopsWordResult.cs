using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Format;

/// <summary>
///     Result for getting tab stops from Word documents.
/// </summary>
public sealed record GetTabStopsWordResult
{
    /// <summary>
    ///     Location type (body, header, footer).
    /// </summary>
    [JsonPropertyName("location")]
    public required string Location { get; init; }

    /// <summary>
    ///     Description of the location.
    /// </summary>
    [JsonPropertyName("locationDescription")]
    public required string LocationDescription { get; init; }

    /// <summary>
    ///     Zero-based section index.
    /// </summary>
    [JsonPropertyName("sectionIndex")]
    public required int SectionIndex { get; init; }

    /// <summary>
    ///     Zero-based paragraph index (only for body location with single paragraph).
    /// </summary>
    [JsonPropertyName("paragraphIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ParagraphIndex { get; init; }

    /// <summary>
    ///     Indicates whether all paragraphs were queried.
    /// </summary>
    [JsonPropertyName("allParagraphs")]
    public required bool AllParagraphs { get; init; }

    /// <summary>
    ///     Indicates whether style tab stops are included.
    /// </summary>
    [JsonPropertyName("includeStyle")]
    public required bool IncludeStyle { get; init; }

    /// <summary>
    ///     Number of paragraphs queried.
    /// </summary>
    [JsonPropertyName("paragraphCount")]
    public required int ParagraphCount { get; init; }

    /// <summary>
    ///     Total number of tab stops found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of tab stops.
    /// </summary>
    [JsonPropertyName("tabStops")]
    public required IReadOnlyList<TabStopDetailInfo> TabStops { get; init; }
}
