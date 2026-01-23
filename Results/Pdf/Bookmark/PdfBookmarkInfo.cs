using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Bookmark;

/// <summary>
///     Information about a single PDF bookmark.
/// </summary>
public record PdfBookmarkInfo
{
    /// <summary>
    ///     Index of the bookmark.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Bookmark title.
    /// </summary>
    [JsonPropertyName("title")]
    public required string Title { get; init; }

    /// <summary>
    ///     Target page index.
    /// </summary>
    [JsonPropertyName("pageIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? PageIndex { get; init; }
}
