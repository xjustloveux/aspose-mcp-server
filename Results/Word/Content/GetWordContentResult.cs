using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Content;

/// <summary>
///     Result for getting Word document content as plain text with pagination.
/// </summary>
public sealed record GetWordContentResult
{
    /// <summary>
    ///     The document content text.
    /// </summary>
    [JsonPropertyName("content")]
    public required string Content { get; init; }

    /// <summary>
    ///     Total length of the document content in characters.
    /// </summary>
    [JsonPropertyName("totalLength")]
    public required int TotalLength { get; init; }

    /// <summary>
    ///     The character offset from the beginning of the document.
    /// </summary>
    [JsonPropertyName("offset")]
    public required int Offset { get; init; }

    /// <summary>
    ///     Indicates whether there is more content available beyond the current page.
    /// </summary>
    [JsonPropertyName("hasMore")]
    public required bool HasMore { get; init; }
}
