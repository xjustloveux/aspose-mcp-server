using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Sheet;

/// <summary>
///     Result for getting worksheets from Excel workbooks.
/// </summary>
public record GetSheetsResult
{
    /// <summary>
    ///     Number of worksheets.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Name of the workbook.
    /// </summary>
    [JsonPropertyName("workbookName")]
    public required string WorkbookName { get; init; }

    /// <summary>
    ///     List of worksheet information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ExcelSheetInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no worksheets found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
