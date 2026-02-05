using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Toc;

/// <summary>
///     Information about a single table of contents entry in a PDF document.
/// </summary>
public record TocEntryPdfInfo
{
    /// <summary>
    ///     Title of the TOC entry.
    /// </summary>
    [JsonPropertyName("title")]
    public required string Title { get; init; }

    /// <summary>
    ///     Target page number (1-based).
    /// </summary>
    [JsonPropertyName("pageNumber")]
    public required int PageNumber { get; init; }

    /// <summary>
    ///     Heading level of the TOC entry (1-based).
    /// </summary>
    [JsonPropertyName("level")]
    public required int Level { get; init; }
}
