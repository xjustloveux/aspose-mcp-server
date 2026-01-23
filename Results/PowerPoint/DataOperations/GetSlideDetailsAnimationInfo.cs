using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Animation information for a slide.
/// </summary>
public sealed record GetSlideDetailsAnimationInfo
{
    /// <summary>
    ///     Gets the animation index.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

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
