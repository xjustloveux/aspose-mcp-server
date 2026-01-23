using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Filter;

/// <summary>
///     Information about a filter column.
/// </summary>
public record FilterColumnInfo
{
    /// <summary>
    ///     Zero-based column index.
    /// </summary>
    [JsonPropertyName("columnIndex")]
    public required int ColumnIndex { get; init; }

    /// <summary>
    ///     Type of filter applied.
    /// </summary>
    [JsonPropertyName("filterType")]
    public required string FilterType { get; init; }

    /// <summary>
    ///     Whether the filter dropdown is visible.
    /// </summary>
    [JsonPropertyName("isDropdownVisible")]
    public required bool IsDropdownVisible { get; init; }
}
