using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Sparkline;

/// <summary>
///     Information about a sparkline group.
/// </summary>
public record ExcelSparklineInfo
{
    /// <summary>
    ///     Sparkline group index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Sparkline type (Line, Column, Stacked).
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Number of sparklines in the group.
    /// </summary>
    [JsonPropertyName("sparklineCount")]
    public int SparklineCount { get; init; }

    /// <summary>
    ///     Data range.
    /// </summary>
    [JsonPropertyName("dataRange")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? DataRange { get; init; }

    /// <summary>
    ///     Location range.
    /// </summary>
    [JsonPropertyName("locationRange")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LocationRange { get; init; }
}
