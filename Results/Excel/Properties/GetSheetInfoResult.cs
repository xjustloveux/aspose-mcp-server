using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Properties;

/// <summary>
///     Result for getting sheet information from Excel files.
/// </summary>
public record GetSheetInfoResult
{
    /// <summary>
    ///     Number of worksheets returned in this result.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Total number of worksheets in the workbook.
    /// </summary>
    [JsonPropertyName("totalWorksheets")]
    public required int TotalWorksheets { get; init; }

    /// <summary>
    ///     List of sheet information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<SheetInfoDetail> Items { get; init; }
}
