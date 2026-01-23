using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Style;

/// <summary>
///     Result for getting cell format information from Excel workbooks.
/// </summary>
public record GetCellFormatResult
{
    /// <summary>
    ///     Number of cells in the result.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     The cell or range that was queried.
    /// </summary>
    [JsonPropertyName("range")]
    public required string Range { get; init; }

    /// <summary>
    ///     The fields that were requested.
    /// </summary>
    [JsonPropertyName("fields")]
    public required string Fields { get; init; }

    /// <summary>
    ///     List of cell format data.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<CellFormatInfo> Items { get; init; }
}
