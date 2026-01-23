using System.Text.Json.Serialization;

namespace AsposeMcpServer.Results.PowerPoint.DataOperations;

/// <summary>
///     Transition information for a slide.
/// </summary>
public sealed record GetSlideDetailsTransitionInfo
{
    /// <summary>
    ///     Gets the transition type.
    /// </summary>
    [JsonPropertyName("type")]
    public required string Type { get; init; }

    /// <summary>
    ///     Gets the transition speed.
    /// </summary>
    [JsonPropertyName("speed")]
    public required string Speed { get; init; }

    /// <summary>
    ///     Gets whether to advance on click.
    /// </summary>
    [JsonPropertyName("advanceOnClick")]
    public required bool AdvanceOnClick { get; init; }

    /// <summary>
    ///     Gets the advance after time in milliseconds.
    /// </summary>
    [JsonPropertyName("advanceAfterTimeMs")]
    public required uint AdvanceAfterTimeMs { get; init; }
}
