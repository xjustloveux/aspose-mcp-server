using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Page;

/// <summary>
///     Basic information about a single PDF page.
/// </summary>
public record PdfPageInfo
{
    /// <summary>
    ///     Page index (1-based).
    /// </summary>
    [JsonPropertyName("pageIndex")]
    public required int PageIndex { get; init; }

    /// <summary>
    ///     Page width.
    /// </summary>
    [JsonPropertyName("width")]
    public required double Width { get; init; }

    /// <summary>
    ///     Page height.
    /// </summary>
    [JsonPropertyName("height")]
    public required double Height { get; init; }

    /// <summary>
    ///     Page rotation.
    /// </summary>
    [JsonPropertyName("rotation")]
    public required string Rotation { get; init; }
}
