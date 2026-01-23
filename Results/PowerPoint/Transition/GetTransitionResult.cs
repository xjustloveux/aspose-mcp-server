using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.Transition;

/// <summary>
///     Result for getting transition from PowerPoint presentations.
/// </summary>
public record GetTransitionResult
{
    /// <summary>
    ///     Slide index.
    /// </summary>
    [JsonPropertyName("slideIndex")]
    public required int SlideIndex { get; init; }

    /// <summary>
    ///     Transition type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Whether slide has transition.
    /// </summary>
    [JsonPropertyName("hasTransition")]
    public required bool HasTransition { get; init; }

    /// <summary>
    ///     Transition speed.
    /// </summary>
    [JsonPropertyName("speed")]
    public required string Speed { get; init; }

    /// <summary>
    ///     Advance on click.
    /// </summary>
    [JsonPropertyName("advanceOnClick")]
    public required bool AdvanceOnClick { get; init; }

    /// <summary>
    ///     Advance after.
    /// </summary>
    [JsonPropertyName("advanceAfter")]
    public required bool AdvanceAfter { get; init; }

    /// <summary>
    ///     Advance after seconds.
    /// </summary>
    [JsonPropertyName("advanceAfterSeconds")]
    public required double AdvanceAfterSeconds { get; init; }
}
