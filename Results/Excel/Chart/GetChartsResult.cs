using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Chart;

/// <summary>
///     Result for getting charts from Excel workbooks.
/// </summary>
public record GetChartsResult
{
    /// <summary>
    ///     Number of charts.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     List of chart information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ChartInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no charts found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
