using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.MergeCells;

/// <summary>
///     Result for getting merged cells from Excel workbooks.
/// </summary>
public record GetMergedCellsResult
{
    /// <summary>
    ///     Number of merged cell ranges.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     List of merged cell information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<MergedCellInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no merged cells found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
