using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Table;

/// <summary>
///     Information about a single Excel table (ListObject).
/// </summary>
public record ExcelTableInfo
{
    /// <summary>
    ///     Table index in the worksheet.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Table name.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Data range address (e.g., "A1:D10").
    /// </summary>
    [JsonPropertyName("range")]
    public required string Range { get; init; }

    /// <summary>
    ///     Whether the table has a header row.
    /// </summary>
    [JsonPropertyName("showHeaderRow")]
    public bool ShowHeaderRow { get; init; }

    /// <summary>
    ///     Whether the table has a totals row.
    /// </summary>
    [JsonPropertyName("showTotals")]
    public bool ShowTotals { get; init; }

    /// <summary>
    ///     Table style name.
    /// </summary>
    [JsonPropertyName("styleName")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? StyleName { get; init; }

    /// <summary>
    ///     Number of data rows (excluding header and totals).
    /// </summary>
    [JsonPropertyName("dataRowCount")]
    public int DataRowCount { get; init; }

    /// <summary>
    ///     Number of columns.
    /// </summary>
    [JsonPropertyName("columnCount")]
    public int ColumnCount { get; init; }
}
