using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.HeaderFooter;

/// <summary>
///     Result for getting headers and footers from Word documents.
/// </summary>
public record GetHeadersFootersResult
{
    /// <summary>
    ///     Total number of sections in document.
    /// </summary>
    [JsonPropertyName("totalSections")]
    public required int TotalSections { get; init; }

    /// <summary>
    ///     Queried section index.
    /// </summary>
    [JsonPropertyName("queriedSectionIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? QueriedSectionIndex { get; init; }

    /// <summary>
    ///     List of section information.
    /// </summary>
    [JsonPropertyName("sections")]
    public required IReadOnlyList<SectionHeaderFooterInfo> Sections { get; init; }
}
