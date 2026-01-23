using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Animation;

/// <summary>
///     Information about a single animation.
/// </summary>
public record AnimationInfo
{
    /// <summary>
    ///     Zero-based index of the animation.
    /// </summary>
    [JsonPropertyName("index")]
    public required int Index { get; init; }

    /// <summary>
    ///     Shape index.
    /// </summary>
    [JsonPropertyName("shapeIndex")]
    public required int ShapeIndex { get; init; }

    /// <summary>
    ///     Shape name.
    /// </summary>
    [JsonPropertyName("shapeName")]
    public required string ShapeName { get; init; }

    /// <summary>
    ///     Effect type.
    /// </summary>
    [JsonPropertyName("effectType")]
    public required string EffectType { get; init; }

    /// <summary>
    ///     Effect subtype.
    /// </summary>
    [JsonPropertyName("effectSubtype")]
    public required string EffectSubtype { get; init; }

    /// <summary>
    ///     Trigger type.
    /// </summary>
    [JsonPropertyName("triggerType")]
    public required string TriggerType { get; init; }

    /// <summary>
    ///     Duration.
    /// </summary>
    [JsonPropertyName("duration")]
    public required float Duration { get; init; }

    /// <summary>
    ///     Delay.
    /// </summary>
    [JsonPropertyName("delay")]
    public required float Delay { get; init; }
}
