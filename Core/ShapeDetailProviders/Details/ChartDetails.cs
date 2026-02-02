using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Detail information for Chart elements.
/// </summary>
public sealed record ChartDetails : ShapeDetails
{
    /// <summary>
    ///     Gets the chart type.
    /// </summary>
    [JsonPropertyName("chartType")]
    public required string ChartType { get; init; }

    /// <summary>
    ///     Gets the chart title text.
    /// </summary>
    [JsonPropertyName("title")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; init; }

    /// <summary>
    ///     Gets whether the chart has a title.
    /// </summary>
    [JsonPropertyName("hasTitle")]
    public required bool HasTitle { get; init; }

    /// <summary>
    ///     Gets whether the chart has a legend.
    /// </summary>
    [JsonPropertyName("hasLegend")]
    public required bool HasLegend { get; init; }

    /// <summary>
    ///     Gets the legend position (e.g., Top, Bottom, Left, Right). Only present when HasLegend is true.
    /// </summary>
    [JsonPropertyName("legendPosition")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LegendPosition { get; init; }

    /// <summary>
    ///     Gets the number of data series.
    /// </summary>
    [JsonPropertyName("seriesCount")]
    public required int SeriesCount { get; init; }

    /// <summary>
    ///     Gets the data series information.
    /// </summary>
    [JsonPropertyName("series")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<ChartSeriesInfo>? Series { get; init; }

    /// <summary>
    ///     Gets the number of categories.
    /// </summary>
    [JsonPropertyName("categoryCount")]
    public required int CategoryCount { get; init; }

    /// <summary>
    ///     Gets whether the chart has a data table.
    /// </summary>
    [JsonPropertyName("hasDataTable")]
    public required bool HasDataTable { get; init; }
}

/// <summary>
///     Information about a chart data series.
/// </summary>
public sealed record ChartSeriesInfo
{
    /// <summary>
    ///     Gets the series name.
    /// </summary>
    [JsonPropertyName("name")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Name { get; init; }

    /// <summary>
    ///     Gets the series type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }
}
