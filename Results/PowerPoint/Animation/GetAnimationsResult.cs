using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Animation;

/// <summary>
///     Result for getting animations from PowerPoint presentations.
/// </summary>
public record GetAnimationsResult
{
    /// <summary>
    ///     Slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Shape index filter if specified.
    /// </summary>
    [JsonPropertyName("filterByShapeIndex")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? FilterByShapeIndex { get; init; }

    /// <summary>
    ///     Total number of animations on the slide.
    /// </summary>
    [JsonPropertyName("totalAnimationsOnSlide")]
    public required int TotalAnimationsOnSlide { get; init; }

    /// <summary>
    ///     List of animation information.
    /// </summary>
    [JsonPropertyName("animations")]
    public required IReadOnlyList<AnimationInfo> Animations { get; init; }
}
