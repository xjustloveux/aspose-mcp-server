using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Animation information for a slide.
/// </summary>
public sealed record GetSlideDetailsAnimationInfo
{
    /// <summary>
    ///     Animation index relative to its target shape (0-based within that shape) — pass together with
    ///     shapeIndex to ppt_animation edit/delete to address this exact animation.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Index of the shape this animation targets. Pass with index to ppt_animation edit/delete.
    /// </summary>
    [JsonPropertyName("shapeIndex")]
    public int ShapeIndex { get; init; }

    /// <summary>
    ///     Gets the animation type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Gets the target shape type.
    /// </summary>
    [JsonPropertyName("targetShape")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? TargetShape { get; init; }
}
