using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Watermark;

/// <summary>
///     Information about a watermark on a slide.
/// </summary>
public record PptWatermarkInfo
{
    /// <summary>
    ///     The slide index containing the watermark.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     The shape name of the watermark.
    /// </summary>
    [JsonPropertyName("shapeName")]
    public required string ShapeName { get; init; }

    /// <summary>
    ///     The type of watermark (text or image).
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     The watermark text (for text watermarks).
    /// </summary>
    [JsonPropertyName("text")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Text { get; init; }
}
