using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Word.Content;

/// <summary>
///     Result for getting Word document metadata and properties.
/// </summary>
public sealed record GetWordDocumentInfoResult
{
    /// <summary>
    ///     Document title.
    /// </summary>
    [JsonPropertyName("title")]
    public string? Title { get; init; }

    /// <summary>
    ///     Document author.
    /// </summary>
    [JsonPropertyName("author")]
    public string? Author { get; init; }

    /// <summary>
    ///     Document subject.
    /// </summary>
    [JsonPropertyName("subject")]
    public string? Subject { get; init; }

    /// <summary>
    ///     Document creation date formatted as yyyy-MM-dd HH:mm:ss.
    /// </summary>
    [JsonPropertyName("created")]
    public required string Created { get; init; }

    /// <summary>
    ///     Document last modified date formatted as yyyy-MM-dd HH:mm:ss.
    /// </summary>
    [JsonPropertyName("modified")]
    public required string Modified { get; init; }

    /// <summary>
    ///     Number of pages in the document.
    /// </summary>
    [JsonPropertyName("pages")]
    public required int Pages { get; init; }

    /// <summary>
    ///     Number of sections in the document.
    /// </summary>
    [JsonPropertyName("sections")]
    public required int Sections { get; init; }

    /// <summary>
    ///     Indicates whether tab stops information is included.
    /// </summary>
    [JsonPropertyName("tabStopsIncluded")]
    public required bool TabStopsIncluded { get; init; }

    /// <summary>
    ///     List of tab stops by paragraph when tabStopsIncluded is true.
    /// </summary>
    [JsonPropertyName("tabStops")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<TabStopsByParagraph>? TabStops { get; init; }
}
