using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Result type for getting content from PowerPoint presentations.
/// </summary>
public sealed record GetContentPptResult
{
    /// <summary>
    ///     Gets the total number of slides.
    /// </summary>
    [JsonPropertyName("totalSlides")]
    public required int TotalSlides { get; init; }

    /// <summary>
    ///     Gets the slides with their content.
    /// </summary>
    [JsonPropertyName("slides")]
    public required List<GetContentSlideInfo> Slides { get; init; }
}
