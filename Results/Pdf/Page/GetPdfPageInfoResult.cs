using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Page;

/// <summary>
///     Result for getting page information from PDF documents.
/// </summary>
public record GetPdfPageInfoResult
{
    /// <summary>
    ///     Total number of pages.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of page information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<PdfPageInfo> Items { get; init; }
}
