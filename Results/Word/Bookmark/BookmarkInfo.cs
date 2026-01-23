using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Bookmark;

/// <summary>
///     Information about a single bookmark.
/// </summary>
public record BookmarkInfo
{
    /// <summary>
    ///     Zero-based index of the bookmark.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Name of the bookmark.
    /// </summary>
    [JsonPropertyName("name")]
    public required string Name { get; init; }

    /// <summary>
    ///     Text content within the bookmark.
    /// </summary>
    [JsonPropertyName("text")]
    public required string Text { get; init; }

    /// <summary>
    ///     Length of the bookmark text.
    /// </summary>
    [JsonPropertyName("length")]
    public required int Length { get; init; }
}
