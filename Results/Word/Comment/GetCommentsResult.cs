using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Comment;

/// <summary>
///     Result for getting comments from Word documents.
/// </summary>
public record GetCommentsResult
{
    /// <summary>
    ///     Total number of top-level comments in the document.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of comment information.
    /// </summary>
    [JsonPropertyName("comments")]
    public required IReadOnlyList<CommentInfo> Comments { get; init; }

    /// <summary>
    ///     Optional message when no comments found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
