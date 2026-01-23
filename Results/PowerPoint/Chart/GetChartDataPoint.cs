using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Chart;

/// <summary>
///     A single data point.
/// </summary>
public sealed record GetChartDataPoint
{
    /// <summary>
    ///     Gets the data point index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets the data point value.
    /// </summary>
    [JsonPropertyName("value")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; init; }
}
