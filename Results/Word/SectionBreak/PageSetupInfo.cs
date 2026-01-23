using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.SectionBreak;

/// <summary>
///     Page setup information for a section.
/// </summary>
public record PageSetupInfo
{
    /// <summary>
    ///     Paper size.
    /// </summary>
    [JsonPropertyName("paperSize")]
    public required string PaperSize { get; init; }

    /// <summary>
    ///     Page orientation.
    /// </summary>
    [JsonPropertyName("orientation")]
    public required string Orientation { get; init; }

    /// <summary>
    ///     Page margins.
    /// </summary>
    [JsonPropertyName("margins")]
    public required MarginsInfo Margins { get; init; }

    /// <summary>
    ///     Header and footer distance from page edge.
    /// </summary>
    [JsonPropertyName("headerFooterDistance")]
    public required HeaderFooterDistanceInfo HeaderFooterDistance { get; init; }

    /// <summary>
    ///     Page number start (null if not restarted).
    /// </summary>
    [JsonPropertyName("pageNumberStart")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? PageNumberStart { get; init; }

    /// <summary>
    ///     Whether the first page has different header/footer.
    /// </summary>
    [JsonPropertyName("differentFirstPage")]
    public required bool DifferentFirstPage { get; init; }

    /// <summary>
    ///     Whether odd and even pages have different headers/footers.
    /// </summary>
    [JsonPropertyName("differentOddEvenPages")]
    public required bool DifferentOddEvenPages { get; init; }

    /// <summary>
    ///     Number of text columns.
    /// </summary>
    [JsonPropertyName("columnCount")]
    public required int ColumnCount { get; init; }
}
