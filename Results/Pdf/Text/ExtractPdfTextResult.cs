using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Text;

/// <summary>
///     Result for extracting text from PDF documents.
/// </summary>
public record ExtractPdfTextResult
{
    /// <summary>
    ///     Page index (1-based).
    /// </summary>
    [JsonPropertyName("pageIndex")]
    public required int PageIndex { get; init; }

    /// <summary>
    ///     Total number of pages in the document.
    /// </summary>
    [JsonPropertyName("totalPages")]
    public required int TotalPages { get; init; }

    /// <summary>
    ///     Extracted text content (when includeFontInfo is false).
    /// </summary>
    [JsonPropertyName("text")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; init; }

    /// <summary>
    ///     Number of text fragments (when includeFontInfo is true).
    /// </summary>
    [JsonPropertyName("fragmentCount")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? FragmentCount { get; init; }

    /// <summary>
    ///     List of text fragments with font information (when includeFontInfo is true).
    /// </summary>
    [JsonPropertyName("fragments")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public IReadOnlyList<PdfTextFragment>? Fragments { get; init; }
}
