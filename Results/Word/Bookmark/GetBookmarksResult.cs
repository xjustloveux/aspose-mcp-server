using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Bookmark;

/// <summary>
///     Result for getting bookmarks from Word documents.
/// </summary>
public record GetBookmarksResult
{
    /// <summary>
    ///     Total number of bookmarks in the document.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of bookmark information.
    /// </summary>
    [JsonPropertyName("bookmarks")]
    public required IReadOnlyList<BookmarkInfo> Bookmarks { get; init; }

    /// <summary>
    ///     Optional message when no bookmarks found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
