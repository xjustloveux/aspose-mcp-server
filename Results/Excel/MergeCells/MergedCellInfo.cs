using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.MergeCells;

/// <summary>
///     Information about a single merged cell range.
/// </summary>
public record MergedCellInfo
{
    /// <summary>
    ///     Zero-based index of the merged cell range.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Range string.
    /// </summary>
    [JsonPropertyName("range")]
    public required string Range { get; init; }

    /// <summary>
    ///     Start cell reference.
    /// </summary>
    [JsonPropertyName("startCell")]
    public required string StartCell { get; init; }

    /// <summary>
    ///     End cell reference.
    /// </summary>
    [JsonPropertyName("endCell")]
    public required string EndCell { get; init; }

    /// <summary>
    ///     Number of rows in the merged range.
    /// </summary>
    [JsonPropertyName("rowCount")]
    public required int RowCount { get; init; }

    /// <summary>
    ///     Number of columns in the merged range.
    /// </summary>
    [JsonPropertyName("columnCount")]
    public required int ColumnCount { get; init; }

    /// <summary>
    ///     Cell value.
    /// </summary>
    [JsonPropertyName("value")]
    public required string Value { get; init; }
}
