using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Range;

/// <summary>
///     Result for getting data from Excel ranges.
/// </summary>
public record GetRangeResult
{
    /// <summary>
    ///     The range address that was queried.
    /// </summary>
    [JsonPropertyName("range")]
    public required string Range { get; init; }

    /// <summary>
    ///     Number of rows in the range.
    /// </summary>
    [JsonPropertyName("rowCount")]
    public required int RowCount { get; init; }

    /// <summary>
    ///     Number of columns in the range.
    /// </summary>
    [JsonPropertyName("columnCount")]
    public required int ColumnCount { get; init; }

    /// <summary>
    ///     Total number of cells in the range.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of cell data.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<RangeCellInfo> Items { get; init; }
}
