using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Info;

/// <summary>
///     Result for getting text content from PDF documents.
/// </summary>
public record GetPdfContentResult
{
    /// <summary>
    ///     Page index if extracting from a specific page.
    /// </summary>
    [JsonPropertyName("pageIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? PageIndex { get; init; }

    /// <summary>
    ///     Total number of pages in the document.
    /// </summary>
    [JsonPropertyName("totalPages")]
    public required int TotalPages { get; init; }

    /// <summary>
    ///     Number of pages that were extracted.
    /// </summary>
    [JsonPropertyName("extractedPages")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? ExtractedPages { get; init; }

    /// <summary>
    ///     Whether the content was truncated due to maxPages limit.
    /// </summary>
    [JsonPropertyName("truncated")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? Truncated { get; init; }

    /// <summary>
    ///     The extracted text content.
    /// </summary>
    [JsonPropertyName("content")]
    public required string Content { get; init; }
}
