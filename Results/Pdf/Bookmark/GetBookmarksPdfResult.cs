using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Bookmark;

/// <summary>
///     Result for getting bookmarks from PDF documents.
/// </summary>
public record GetBookmarksPdfResult
{
    /// <summary>
    ///     Number of bookmarks.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of bookmark information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<PdfBookmarkInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no bookmarks found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
