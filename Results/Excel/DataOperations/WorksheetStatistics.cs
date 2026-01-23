using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataOperations;

/// <summary>
///     Statistics for a single worksheet.
/// </summary>
public record WorksheetStatistics
{
    /// <summary>
    ///     Zero-based index of the worksheet.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Maximum row containing data (1-based count).
    /// </summary>
    [JsonPropertyName("maxDataRow")]
    public required int MaxDataRow { get; init; }

    /// <summary>
    ///     Maximum column containing data (1-based count).
    /// </summary>
    [JsonPropertyName("maxDataColumn")]
    public required int MaxDataColumn { get; init; }

    /// <summary>
    ///     Number of charts in the worksheet.
    /// </summary>
    [JsonPropertyName("chartsCount")]
    public required int ChartsCount { get; init; }

    /// <summary>
    ///     Number of pictures in the worksheet.
    /// </summary>
    [JsonPropertyName("picturesCount")]
    public required int PicturesCount { get; init; }

    /// <summary>
    ///     Number of hyperlinks in the worksheet.
    /// </summary>
    [JsonPropertyName("hyperlinksCount")]
    public required int HyperlinksCount { get; init; }

    /// <summary>
    ///     Number of comments in the worksheet.
    /// </summary>
    [JsonPropertyName("commentsCount")]
    public required int CommentsCount { get; init; }

    /// <summary>
    ///     Range statistics (only present when a range is specified).
    /// </summary>
    [JsonPropertyName("rangeStatistics")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public RangeStatistics? RangeStatistics { get; init; }

    /// <summary>
    ///     Error message when range statistics calculation fails.
    /// </summary>
    [JsonPropertyName("rangeStatisticsError")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RangeStatisticsError { get; init; }
}
