using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.HeaderFooter;

/// <summary>
///     Result for PDF header/footer operations.
/// </summary>
public record HeaderFooterPdfResult
{
    /// <summary>
    ///     The number of pages affected by the operation.
    /// </summary>
    [JsonPropertyName("pagesAffected")]
    public required int PagesAffected { get; init; }

    /// <summary>
    ///     The position of the header or footer ("header" or "footer").
    /// </summary>
    [JsonPropertyName("position")]
    public required string Position { get; init; }

    /// <summary>
    ///     Human-readable message describing the operation result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
