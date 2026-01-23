using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Properties;

/// <summary>
///     Information about worksheet print settings.
/// </summary>
public record PrintSettingsInfo
{
    /// <summary>
    ///     The print area range.
    /// </summary>
    [JsonPropertyName("printArea")]
    public string? PrintArea { get; init; }

    /// <summary>
    ///     Rows to repeat at top when printing.
    /// </summary>
    [JsonPropertyName("printTitleRows")]
    public string? PrintTitleRows { get; init; }

    /// <summary>
    ///     Columns to repeat at left when printing.
    /// </summary>
    [JsonPropertyName("printTitleColumns")]
    public string? PrintTitleColumns { get; init; }

    /// <summary>
    ///     Page orientation (Portrait or Landscape).
    /// </summary>
    [JsonPropertyName("orientation")]
    public required string Orientation { get; init; }

    /// <summary>
    ///     Paper size setting.
    /// </summary>
    [JsonPropertyName("paperSize")]
    public required string PaperSize { get; init; }

    /// <summary>
    ///     Number of pages wide to fit when printing.
    /// </summary>
    [JsonPropertyName("fitToPagesWide")]
    public required int FitToPagesWide { get; init; }

    /// <summary>
    ///     Number of pages tall to fit when printing.
    /// </summary>
    [JsonPropertyName("fitToPagesTall")]
    public required int FitToPagesTall { get; init; }
}
