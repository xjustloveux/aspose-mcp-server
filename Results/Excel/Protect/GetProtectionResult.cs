using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Protect;

/// <summary>
///     Result for getting protection status from Excel workbooks.
/// </summary>
public record GetProtectionResult
{
    /// <summary>
    ///     Number of worksheets in result.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Total number of worksheets.
    /// </summary>
    [JsonPropertyName("totalWorksheets")]
    public required int TotalWorksheets { get; init; }

    /// <summary>
    ///     List of worksheet protection information.
    /// </summary>
    [JsonPropertyName("worksheets")]
    public required IReadOnlyList<WorksheetProtectionInfo> Worksheets { get; init; }
}
