using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Page;

/// <summary>
///     Result for getting detailed page information from PDF documents.
/// </summary>
public record GetPdfPageDetailsResult
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

    /// <summary>
    ///     Media box dimensions.
    /// </summary>
    [JsonPropertyName("mediaBox")]
    public required PdfPageBox MediaBox { get; init; }

    /// <summary>
    ///     Crop box dimensions.
    /// </summary>
    [JsonPropertyName("cropBox")]
    public required PdfPageBox CropBox { get; init; }

    /// <summary>
    ///     Number of annotations on the page.
    /// </summary>
    [JsonPropertyName("annotations")]
    public required int Annotations { get; init; }

    /// <summary>
    ///     Number of paragraphs on the page.
    /// </summary>
    [JsonPropertyName("paragraphs")]
    public required int Paragraphs { get; init; }

    /// <summary>
    ///     Number of images on the page.
    /// </summary>
    [JsonPropertyName("images")]
    public required int Images { get; init; }
}
