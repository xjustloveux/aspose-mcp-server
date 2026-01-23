using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Sheet;

/// <summary>
///     Information about a single worksheet.
/// </summary>
public record ExcelSheetInfo
{
    /// <summary>
    ///     Zero-based index of the worksheet.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Visibility status.
    /// </summary>
    [JsonPropertyName("visibility")]
    public required string Visibility { get; init; }
}
