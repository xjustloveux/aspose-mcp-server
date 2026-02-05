using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Table;

/// <summary>
///     Result for getting tables from an Excel worksheet.
/// </summary>
public record GetTablesExcelResult
{
    /// <summary>
    ///     Number of tables found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Sheet index.
    /// </summary>
    [JsonPropertyName("sheetIndex")]
    public required int SheetIndex { get; init; }

    /// <summary>
    ///     List of table information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ExcelTableInfo> Items { get; init; }

    /// <summary>
    ///     Optional message.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
