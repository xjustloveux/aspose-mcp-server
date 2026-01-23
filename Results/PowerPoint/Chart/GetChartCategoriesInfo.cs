using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Chart;

/// <summary>
///     Categories information for chart data.
/// </summary>
public sealed record GetChartCategoriesInfo
{
    /// <summary>
    ///     Gets the count of categories.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Gets the category items.
    /// </summary>
    [JsonPropertyName("items")]
    public required List<GetChartCategoryItem> Items { get; init; }
}
