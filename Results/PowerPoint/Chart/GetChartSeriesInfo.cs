using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Chart;

/// <summary>
///     Series information for chart data.
/// </summary>
public sealed record GetChartSeriesInfo
{
    /// <summary>
    ///     Gets the count of series.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Gets the series items.
    /// </summary>
    [JsonPropertyName("items")]
    public required List<GetChartSeriesItem> Items { get; init; }
}
