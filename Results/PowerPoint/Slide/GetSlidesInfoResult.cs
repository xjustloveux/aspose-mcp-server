using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Slide;

/// <summary>
///     Result type for getting slides information from PowerPoint presentations.
/// </summary>
public sealed record GetSlidesInfoResult
{
    /// <summary>
    ///     Gets the count of slides.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     Gets the slides list.
    /// </summary>
    [JsonPropertyName("slides")]
    public required List<GetSlideInfoItem> Slides { get; init; }

    /// <summary>
    ///     Gets the available layouts.
    /// </summary>
    [JsonPropertyName("availableLayouts")]
    public required List<GetSlideLayoutInfo> AvailableLayouts { get; init; }
}
