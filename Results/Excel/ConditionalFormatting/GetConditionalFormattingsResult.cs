using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.ConditionalFormatting;

/// <summary>
///     Result for getting conditional formattings from Excel workbooks.
/// </summary>
public record GetConditionalFormattingsResult
{
    /// <summary>
    ///     Number of conditional formattings.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Sheet index.
    /// </summary>
    [JsonPropertyName("sheetIndex")]
    public required int SheetIndex { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     List of conditional formatting information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ConditionalFormattingInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no conditional formattings found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
