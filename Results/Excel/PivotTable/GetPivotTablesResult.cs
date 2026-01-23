using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.PivotTable;

/// <summary>
///     Result for getting pivot tables from Excel workbooks.
/// </summary>
public record GetPivotTablesResult
{
    /// <summary>
    ///     Number of pivot tables.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     List of pivot table information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<PivotTableInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no pivot tables found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
