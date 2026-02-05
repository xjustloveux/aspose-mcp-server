using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.PageBreak;

/// <summary>
///     Information about a single page break.
/// </summary>
public record ExcelPageBreakInfo
{
    /// <summary>
    ///     Page break index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Page break type (Horizontal or Vertical).
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Starting row/column index.
    /// </summary>
    [JsonPropertyName("startIndex")]
    public int StartIndex { get; init; }

    /// <summary>
    ///     Ending row/column index.
    /// </summary>
    [JsonPropertyName("endIndex")]
    public int EndIndex { get; init; }
}
