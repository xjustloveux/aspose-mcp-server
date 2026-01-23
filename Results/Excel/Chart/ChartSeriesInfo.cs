using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Chart;

/// <summary>
///     Chart series information.
/// </summary>
public record ChartSeriesInfo
{
    /// <summary>
    ///     Zero-based index of the series.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Series name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Values range.
    /// </summary>
    [JsonPropertyName("valuesRange")]
    public required string ValuesRange { get; init; }

    /// <summary>
    ///     Category data range.
    /// </summary>
    [JsonPropertyName("categoryData")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? CategoryData { get; init; }
}
