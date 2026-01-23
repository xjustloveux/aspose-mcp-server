using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Comment;

/// <summary>
///     Result for getting comments from Excel workbooks.
/// </summary>
public record GetCommentsExcelResult
{
    /// <summary>
    ///     Number of comments found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Sheet index.
    /// </summary>
    [JsonPropertyName("sheetIndex")]
    public required int SheetIndex { get; init; }

    /// <summary>
    ///     Specific cell if queried.
    /// </summary>
    [JsonPropertyName("cell")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Cell { get; init; }

    /// <summary>
    ///     List of comment information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<ExcelCommentInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no comments found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
