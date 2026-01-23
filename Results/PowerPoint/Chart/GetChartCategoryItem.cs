using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Chart;

/// <summary>
///     A single category item.
/// </summary>
public sealed record GetChartCategoryItem
{
    /// <summary>
    ///     Gets the category index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Gets the category value.
    /// </summary>
    [JsonPropertyName("value")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Value { get; init; }
}
