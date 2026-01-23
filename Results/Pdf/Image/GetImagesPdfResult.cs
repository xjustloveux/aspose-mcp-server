using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Image;

/// <summary>
///     Result for getting images from PDF documents.
/// </summary>
public record GetImagesPdfResult
{
    /// <summary>
    ///     Number of images found.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Page index if specified.
    /// </summary>
    [JsonPropertyName("pageIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? PageIndex { get; init; }

    /// <summary>
    ///     List of image information.
    /// </summary>
    [JsonPropertyName("items")]
    public required IReadOnlyList<PdfImageInfo> Items { get; init; }

    /// <summary>
    ///     Optional message when no images found.
    /// </summary>
    [JsonPropertyName("message")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
}
