using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Paragraph;

/// <summary>
///     Result for getting paragraphs from Word documents.
/// </summary>
public record GetParagraphsWordResult
{
    /// <summary>
    ///     Number of paragraphs found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Applied filters.
    /// </summary>
    [JsonPropertyName("filters")]
    public required ParagraphFilters Filters { get; init; }

    /// <summary>
    ///     List of paragraph information.
    /// </summary>
    [JsonPropertyName("paragraphs")]
    public required IReadOnlyList<ParagraphInfo> Paragraphs { get; init; }
}
