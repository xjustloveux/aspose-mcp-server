using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Excel.Comment;

/// <summary>
///     Information about a single Excel comment.
/// </summary>
public record ExcelCommentInfo
{
    /// <summary>
    ///     Cell address.
    /// </summary>
    [JsonPropertyName("cell")]
    public required string Cell { get; init; }

    /// <summary>
    ///     Comment author.
    /// </summary>
    [JsonPropertyName("author")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Author { get; init; }

    /// <summary>
    ///     Comment note/text.
    /// </summary>
    [JsonPropertyName("note")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Note { get; init; }
}
