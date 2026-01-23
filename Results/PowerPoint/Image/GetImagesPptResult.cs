using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Image;

/// <summary>
///     Result for getting images from PowerPoint presentations.
/// </summary>
public record GetImagesPptResult
{
    /// <summary>
    ///     Slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Number of images found.
    /// </summary>
    [JsonPropertyName("imageCount")]
    public required int ImageCount { get; init; }

    /// <summary>
    ///     List of image information.
    /// </summary>
    [JsonPropertyName("images")]
    public required IReadOnlyList<PptImageInfo> Images { get; init; }
}
