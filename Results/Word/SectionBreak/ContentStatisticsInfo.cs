using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.SectionBreak;

/// <summary>
///     Content statistics for a section.
/// </summary>
public record ContentStatisticsInfo
{
    /// <summary>
    ///     Number of paragraphs.
    /// </summary>
    [JsonPropertyName("paragraphs")]
    public required int Paragraphs { get; init; }

    /// <summary>
    ///     Number of tables.
    /// </summary>
    [JsonPropertyName("tables")]
    public required int Tables { get; init; }

    /// <summary>
    ///     Number of shapes.
    /// </summary>
    [JsonPropertyName("shapes")]
    public required int Shapes { get; init; }
}
