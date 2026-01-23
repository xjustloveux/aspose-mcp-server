using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.DataOperations;

/// <summary>
///     Result for getting the used range information from Excel worksheets.
/// </summary>
public record GetUsedRangeResult
{
    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     Zero-based index of the worksheet.
    /// </summary>
    [JsonPropertyName("sheetIndex")]
    public required int SheetIndex { get; init; }

    /// <summary>
    ///     Zero-based index of the first row containing data.
    /// </summary>
    [JsonPropertyName("firstRow")]
    public required int FirstRow { get; init; }

    /// <summary>
    ///     Zero-based index of the last row containing data.
    /// </summary>
    [JsonPropertyName("lastRow")]
    public required int LastRow { get; init; }

    /// <summary>
    ///     Zero-based index of the first column containing data.
    /// </summary>
    [JsonPropertyName("firstColumn")]
    public required int FirstColumn { get; init; }

    /// <summary>
    ///     Zero-based index of the last column containing data.
    /// </summary>
    [JsonPropertyName("lastColumn")]
    public required int LastColumn { get; init; }

    /// <summary>
    ///     The range address in A1 notation (e.g., "A1:D10"), or null if no data.
    /// </summary>
    [JsonPropertyName("range")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Range { get; init; }
}
