using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Sparkline;

/// <summary>
///     Result for getting sparkline groups from an Excel worksheet.
/// </summary>
public record GetSparklinesExcelResult
{
    /// <summary>
    ///     Number of sparkline groups found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Sheet index.
    /// </summary>
    [JsonPropertyName("sheetIndex")]
    public required int SheetIndex { get; init; }

    /// <summary>
    ///     List of sparkline group information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ExcelSparklineInfo> Items { get; init; }

    /// <summary>
    ///     Optional message.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
