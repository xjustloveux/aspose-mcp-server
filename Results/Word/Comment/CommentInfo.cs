using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Comment;

/// <summary>
///     Information about a single comment.
/// </summary>
public record CommentInfo
{
    /// <summary>
    ///     Zero-based index of the comment (-1 for reply comments).
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Author of the comment.
    /// </summary>
    [JsonPropertyName("author")]
    public required string Author { get; init; }

    /// <summary>
    ///     Author initials.
    /// </summary>
    [JsonPropertyName("initial")]
    public required string Initial { get; init; }

    /// <summary>
    ///     Date and time of the comment.
    /// </summary>
    [JsonPropertyName("date")]
    public required string Date { get; init; }

    /// <summary>
    ///     Text content of the comment.
    /// </summary>
    [JsonPropertyName("content")]
    public required string Content { get; init; }

    /// <summary>
    ///     Whether the comment has a range in the document.
    /// </summary>
    [JsonPropertyName("hasRange")]
    public required bool HasRange { get; init; }

    /// <summary>
    ///     Number of replies to this comment.
    /// </summary>
    [JsonPropertyName("replyCount")]
    public required int ReplyCount { get; init; }

    /// <summary>
    ///     List of reply comments.
    /// </summary>
    [JsonPropertyName("replies")]
    public required IReadOnlyList<CommentInfo> Replies { get; init; }
}
