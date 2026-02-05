using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Comment;

/// <summary>
///     Result of getting comments from a presentation.
/// </summary>
public record GetCommentsPptResult
{
    /// <summary>
    ///     Total number of comments found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     The slide index queried.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     The list of comment information.
    /// </summary>
    [JsonPropertyName("items")]
    public required List<PptCommentInfo> Items { get; init; }

    /// <summary>
    ///     Human-readable message describing the result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
