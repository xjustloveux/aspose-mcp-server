using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.Pdf.Image;

/// <summary>
///     Information about a single PDF image.
/// </summary>
public record PdfImageInfo
{
    /// <summary>
    ///     Image index within the page.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Page index.
    /// </summary>
    [JsonPropertyName("pageIndex")]
    public required int PageIndex { get; init; }

    /// <summary>
    ///     Image width.
    /// </summary>
    [JsonPropertyName("width")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Width { get; init; }

    /// <summary>
    ///     Image height.
    /// </summary>
    [JsonPropertyName("height")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Height { get; init; }

    /// <summary>
    ///     Error message if image info retrieval failed.
    /// </summary>
    [JsonPropertyName("error")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Error { get; init; }
}
