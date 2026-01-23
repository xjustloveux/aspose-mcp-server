using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Properties;

/// <summary>
///     Detailed information about a single worksheet.
/// </summary>
public record SheetInfoDetail
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
    ///     Visibility type of the worksheet.
    /// </summary>
    [JsonPropertyName("visibility")]
    public required string Visibility { get; init; }

    /// <summary>
    ///     Number of rows containing data.
    /// </summary>
    [JsonPropertyName("dataRowCount")]
    public required int DataRowCount { get; init; }

    /// <summary>
    ///     Number of columns containing data.
    /// </summary>
    [JsonPropertyName("dataColumnCount")]
    public required int DataColumnCount { get; init; }

    /// <summary>
    ///     Used range information.
    /// </summary>
    [JsonPropertyName("usedRange")]
    public required UsedRangeInfo UsedRange { get; init; }

    /// <summary>
    ///     Page orientation for printing.
    /// </summary>
    [JsonPropertyName("pageOrientation")]
    public required string PageOrientation { get; init; }

    /// <summary>
    ///     Paper size setting.
    /// </summary>
    [JsonPropertyName("paperSize")]
    public required string PaperSize { get; init; }

    /// <summary>
    ///     Freeze panes information.
    /// </summary>
    [JsonPropertyName("freezePanes")]
    public required FreezePanesInfo FreezePanes { get; init; }
}
