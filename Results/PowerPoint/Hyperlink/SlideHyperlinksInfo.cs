using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Hyperlink;

/// <summary>
///     Hyperlinks information for a single slide.
/// </summary>
public record SlideHyperlinksInfo
{
    /// <summary>
    ///     Slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Number of hyperlinks on this slide.
    /// </summary>
    [JsonPropertyName("count")]
    public required int Count { get; init; }

    /// <summary>
    ///     List of hyperlinks on this slide.
    /// </summary>
    [JsonPropertyName("hyperlinks")]
    public required IReadOnlyList<object> Hyperlinks { get; init; }
}
