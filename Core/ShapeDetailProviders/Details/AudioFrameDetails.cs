using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Detail information for AudioFrame elements.
/// </summary>
public sealed record AudioFrameDetails : ShapeDetails
{
    /// <summary>
    ///     Gets the content type of the embedded audio.
    /// </summary>
    [JsonPropertyName("contentType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ContentType { get; init; }

    /// <summary>
    ///     Gets the audio play mode.
    /// </summary>
    [JsonPropertyName("playMode")]
    public required string PlayMode { get; init; }

    /// <summary>
    ///     Gets the audio volume level.
    /// </summary>
    [JsonPropertyName("volume")]
    public required string Volume { get; init; }

    /// <summary>
    ///     Gets whether the audio plays across slides.
    /// </summary>
    [JsonPropertyName("playAcrossSlides")]
    public required bool PlayAcrossSlides { get; init; }

    /// <summary>
    ///     Gets whether the audio rewinds after playing.
    /// </summary>
    [JsonPropertyName("rewindAudio")]
    public required bool RewindAudio { get; init; }

    /// <summary>
    ///     Gets whether the audio frame is hidden during the show.
    /// </summary>
    [JsonPropertyName("hideAtShowing")]
    public required bool HideAtShowing { get; init; }

    /// <summary>
    ///     Gets whether the audio plays in a loop.
    /// </summary>
    [JsonPropertyName("playLoopMode")]
    public required bool PlayLoopMode { get; init; }
}
