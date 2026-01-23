using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Chart;

/// <summary>
///     A single series item.
/// </summary>
public sealed record GetChartSeriesItem
{
    /// <summary>
    ///     Gets the series index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets the series name.
    /// </summary>
    [JsonPropertyName("name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; init; }

    /// <summary>
    ///     Gets the data points count.
    /// </summary>
    [JsonPropertyName("dataPointsCount")]
    public required int DataPointsCount { get; init; }

    /// <summary>
    ///     Gets the data points.
    /// </summary>
    [JsonPropertyName("dataPoints")]
    public required List<GetChartDataPoint> DataPoints { get; init; }
}
