using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.SectionBreak;

/// <summary>
///     Header and footer distance from page edge.
/// </summary>
public record HeaderFooterDistanceInfo
{
    /// <summary>
    ///     Header distance in points.
    /// </summary>
    [JsonPropertyName("header")]
    public required double Header { get; init; }

    /// <summary>
    ///     Footer distance in points.
    /// </summary>
    [JsonPropertyName("footer")]
    public required double Footer { get; init; }
}
