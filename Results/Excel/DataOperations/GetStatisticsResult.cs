using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataOperations;

/// <summary>
///     Result for getting statistics from Excel worksheets.
/// </summary>
public record GetStatisticsResult
{
    /// <summary>
    ///     Total number of worksheets in the workbook.
    /// </summary>
    [JsonPropertyName("totalWorksheets")]
    public required int TotalWorksheets { get; init; }

    /// <summary>
    ///     File format of the workbook.
    /// </summary>
    [JsonPropertyName("fileFormat")]
    public required string FileFormat { get; init; }

    /// <summary>
    ///     List of worksheet statistics.
    /// </summary>
    [JsonPropertyName("worksheets")]
    public required IReadOnlyList<WorksheetStatistics> Worksheets { get; init; }
}
