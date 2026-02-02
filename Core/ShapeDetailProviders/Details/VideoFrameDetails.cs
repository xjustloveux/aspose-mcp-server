using System.Text.Json.Serialization;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Details;

/// <summary>
///     Detail information for VideoFrame elements.
/// </summary>
public sealed record VideoFrameDetails : ShapeDetails
{
    /// <summary>
    ///     Gets the content type of the embedded video.
    /// </summary>
    [JsonPropertyName("contentType")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? ContentType { get; init; }

    /// <summary>
    ///     Gets the video play mode.
    /// </summary>
    [JsonPropertyName("playMode")]
    public required string PlayMode { get; init; }

    /// <summary>
    ///     Gets the video volume level.
    /// </summary>
    [JsonPropertyName("volume")]
    public required string Volume { get; init; }

    /// <summary>
    ///     Gets whether the video plays in full screen mode.
    /// </summary>
    [JsonPropertyName("fullScreenMode")]
    public required bool FullScreenMode { get; init; }

    /// <summary>
    ///     Gets whether the video frame is hidden during the show.
    /// </summary>
    [JsonPropertyName("hideAtShowing")]
    public required bool HideAtShowing { get; init; }

    /// <summary>
    ///     Gets whether the video plays in a loop.
    /// </summary>
    [JsonPropertyName("playLoopMode")]
    public required bool PlayLoopMode { get; init; }

    /// <summary>
    ///     Gets whether the video rewinds after playing.
    /// </summary>
    [JsonPropertyName("rewindVideo")]
    public required bool RewindVideo { get; init; }

    /// <summary>
    ///     Gets the external link path for the video.
    /// </summary>
    [JsonPropertyName("linkPathLong")]
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LinkPathLong { get; init; }
}
