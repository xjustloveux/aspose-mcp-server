using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Hyperlink;

/// <summary>
///     Result for getting hyperlinks from Excel workbooks.
/// </summary>
public record GetHyperlinksExcelResult
{
    /// <summary>
    ///     Number of hyperlinks found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Name of the worksheet.
    /// </summary>
    [JsonPropertyName("worksheetName")]
    public required string WorksheetName { get; init; }

    /// <summary>
    ///     List of hyperlink information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ExcelHyperlinkInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no hyperlinks found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
