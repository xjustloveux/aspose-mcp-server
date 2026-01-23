using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Content;

/// <summary>
///     Result for getting Word document statistics.
/// </summary>
public sealed record GetWordStatisticsResult
{
    /// <summary>
    ///     Number of pages in the document.
    /// </summary>
    [JsonPropertyName("pages")]
    public required int Pages { get; init; }

    /// <summary>
    ///     Number of words in the document.
    /// </summary>
    [JsonPropertyName("words")]
    public required int Words { get; init; }

    /// <summary>
    ///     Number of characters in the document (without spaces).
    /// </summary>
    [JsonPropertyName("characters")]
    public required int Characters { get; init; }

    /// <summary>
    ///     Number of characters in the document (with spaces).
    /// </summary>
    [JsonPropertyName("charactersWithSpaces")]
    public required int CharactersWithSpaces { get; init; }

    /// <summary>
    ///     Number of paragraphs in the document.
    /// </summary>
    [JsonPropertyName("paragraphs")]
    public required int Paragraphs { get; init; }

    /// <summary>
    ///     Number of lines in the document.
    /// </summary>
    [JsonPropertyName("lines")]
    public required int Lines { get; init; }

    /// <summary>
    ///     Number of footnotes in the document (null if footnotesIncluded is false).
    /// </summary>
    [JsonPropertyName("footnotes")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Footnotes { get; init; }

    /// <summary>
    ///     Indicates whether footnotes count is included.
    /// </summary>
    [JsonPropertyName("footnotesIncluded")]
    public required bool FootnotesIncluded { get; init; }

    /// <summary>
    ///     Number of tables in the document.
    /// </summary>
    [JsonPropertyName("tables")]
    public required int Tables { get; init; }

    /// <summary>
    ///     Number of images in the document.
    /// </summary>
    [JsonPropertyName("images")]
    public required int Images { get; init; }

    /// <summary>
    ///     Number of shapes in the document.
    /// </summary>
    [JsonPropertyName("shapes")]
    public required int Shapes { get; init; }

    /// <summary>
    ///     Indicates that statistics were updated before retrieval.
    /// </summary>
    [JsonPropertyName("statisticsUpdated")]
    public required bool StatisticsUpdated { get; init; }
}
