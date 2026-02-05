using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Stamp;

/// <summary>
///     Result for listing stamp annotations from PDF documents.
/// </summary>
public record GetStampsPdfResult
{
    /// <summary>
    ///     The total number of stamp annotations found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     The list of stamp annotation information.
    /// </summary>
    [JsonPropertyName("stamps")]
    public required IReadOnlyList<PdfStampInfo> Stamps { get; init; }

    /// <summary>
    ///     Human-readable message describing the result.
    /// </summary>
    [JsonPropertyName("message")]
    public required string Message { get; init; }
}
