using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Comment;

/// <summary>
///     Information about a PowerPoint comment.
/// </summary>
public record PptCommentInfo
{
    /// <summary>
    ///     The comment index within the slide.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     The author of the comment.
    /// </summary>
    [JsonPropertyName("author")]
    public required string Author { get; init; }

    /// <summary>
    ///     The comment text.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     The X position of the comment.
    /// </summary>
    [JsonPropertyName("x")]
    public required float X { get; init; }

    /// <summary>
    ///     The Y position of the comment.
    /// </summary>
    [JsonPropertyName("y")]
    public required float Y { get; init; }

    /// <summary>
    ///     The creation date of the comment.
    /// </summary>
    [JsonPropertyName("createdTime")]
    public required string CreatedTime { get; init; }

    /// <summary>
    ///     The number of replies to this comment.
    /// </summary>
    [JsonPropertyName("replyCount")]
    public required int ReplyCount { get; init; }
}
