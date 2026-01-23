using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.PivotTable;

/// <summary>
///     Information about a single pivot table.
/// </summary>
public record PivotTableInfo
{
    /// <summary>
    ///     Zero-based index of the pivot table.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Pivot table name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Data source information.
    /// </summary>
    [JsonPropertyName("dataSource")]
    public required string DataSource { get; init; }

    /// <summary>
    ///     Location information.
    /// </summary>
    [JsonPropertyName("location")]
    public required PivotTableLocation Location { get; init; }

    /// <summary>
    ///     Row fields.
    /// </summary>
    [JsonPropertyName("rowFields")]
    public required IReadOnlyList<PivotFieldInfo> RowFields { get; init; }

    /// <summary>
    ///     Column fields.
    /// </summary>
    [JsonPropertyName("columnFields")]
    public required IReadOnlyList<PivotFieldInfo> ColumnFields { get; init; }

    /// <summary>
    ///     Data fields.
    /// </summary>
    [JsonPropertyName("dataFields")]
    public required IReadOnlyList<PivotDataFieldInfo> DataFields { get; init; }
}
