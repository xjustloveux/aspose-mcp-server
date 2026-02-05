using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.PageBreak;

/// <summary>
///     Result for getting page breaks from an Excel worksheet.
/// </summary>
public record GetPageBreaksExcelResult
{
    /// <summary>
    ///     Total number of page breaks found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Sheet index.
    /// </summary>
    [JsonPropertyName("sheetIndex")]
    public required int SheetIndex { get; init; }

    /// <summary>
    ///     List of page break information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ExcelPageBreakInfo> Items { get; init; }

    /// <summary>
    ///     Optional message.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
