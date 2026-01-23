using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Chart;

/// <summary>
///     Information about a single chart.
/// </summary>
public record ChartInfo
{
    /// <summary>
    ///     Zero-based index of the chart.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Chart name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Chart type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Chart location.
    /// </summary>
    [JsonPropertyName("location")]
    public required ChartLocation Location { get; init; }

    /// <summary>
    ///     Chart width.
    /// </summary>
    [JsonPropertyName("width")]
    public required int Width { get; init; }

    /// <summary>
    ///     Chart height.
    /// </summary>
    [JsonPropertyName("height")]
    public required int Height { get; init; }

    /// <summary>
    ///     Chart title.
    /// </summary>
    [JsonPropertyName("title")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Title { get; init; }

    /// <summary>
    ///     Whether legend is enabled.
    /// </summary>
    [JsonPropertyName("legendEnabled")]
    public required bool LegendEnabled { get; init; }

    /// <summary>
    ///     Number of series.
    /// </summary>
    [JsonPropertyName("seriesCount")]
    public required int SeriesCount { get; init; }

    /// <summary>
    ///     List of series information.
    /// </summary>
    [JsonPropertyName("series")]
    public required IReadOnlyList<ChartSeriesInfo> Series { get; init; }
}
