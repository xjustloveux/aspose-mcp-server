using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Chart;

/// <summary>
///     Result type for getting chart data from PowerPoint presentations.
/// </summary>
public sealed record GetChartDataPptResult
{
    /// <summary>
    ///     Gets the slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Gets the chart index.
    /// </summary>
    [JsonPropertyName("chartIndex")]
    public required int ChartIndex { get; init; }

    /// <summary>
    ///     Gets the chart type.
    /// </summary>
    [JsonPropertyName("chartType")]
    public required string ChartType { get; init; }

    /// <summary>
    ///     Gets whether the chart has a title.
    /// </summary>
    [JsonPropertyName("hasTitle")]
    public required bool HasTitle { get; init; }

    /// <summary>
    ///     Gets the chart title.
    /// </summary>
    [JsonPropertyName("title")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; init; }

    /// <summary>
    ///     Gets the categories information.
    /// </summary>
    [JsonPropertyName("categories")]
    public required GetChartCategoriesInfo Categories { get; init; }

    /// <summary>
    ///     Gets the series information.
    /// </summary>
    [JsonPropertyName("series")]
    public required GetChartSeriesInfo Series { get; init; }
}
