using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Filter;

/// <summary>
///     Result for getting filter status from Excel worksheets.
/// </summary>
public record GetFilterStatusResult
{
    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     Whether auto filter is enabled.
    /// </summary>
    [JsonPropertyName("isFilterEnabled")]
    public required bool IsFilterEnabled { get; init; }

    /// <summary>
    ///     Whether there are active filter criteria.
    /// </summary>
    [JsonPropertyName("hasActiveFilters")]
    public required bool HasActiveFilters { get; init; }

    /// <summary>
    ///     Human-readable status description.
    /// </summary>
    [JsonPropertyName("status")]
    public required string Status { get; init; }

    /// <summary>
    ///     Filter range address, or null when filter is not enabled.
    /// </summary>
    [JsonPropertyName("filterRange")]
    public string? FilterRange { get; init; }

    /// <summary>
    ///     Number of filter columns.
    /// </summary>
    [JsonPropertyName("filterColumnsCount")]
    public required int FilterColumnsCount { get; init; }

    /// <summary>
    ///     List of filter column information.
    /// </summary>
    [JsonPropertyName("filterColumns")]
    public required IReadOnlyList<FilterColumnInfo> FilterColumns { get; init; }
}
