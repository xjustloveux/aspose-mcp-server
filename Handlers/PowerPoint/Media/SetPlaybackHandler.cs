using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Media;

/// <summary>
///     Handler for setting media playback options in PowerPoint slides.
/// </summary>
public class SetPlaybackHandler : OperationHandlerBase<Presentation>
{
    private static readonly Dictionary<string, AudioVolumeMode> VolumeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["mute"] = AudioVolumeMode.Mute,
        ["low"] = AudioVolumeMode.Low,
        ["medium"] = AudioVolumeMode.Medium,
        ["loud"] = AudioVolumeMode.Loud
    };

    private static readonly string SupportedVolumes = string.Join(", ", VolumeMap.Keys);

    /// <inheritdoc />
    public override string Operation => "set_playback";

    /// <summary>
    ///     Sets media playback options for audio or video.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex
    ///     Optional: slideIndex, playMode, loop, rewind, volume
    /// </param>
    /// <returns>Success message with playback settings details.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when shapeIndex is not provided, volume is invalid, or shape is not a media frame.
    /// </exception>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractPlaybackParameters(parameters);

        if (!p.ShapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for set_playback operation");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var shape = PowerPointHelper.GetShape(slide, p.ShapeIndex.Value);

        if (!VolumeMap.TryGetValue(p.Volume, out var volume))
            throw new ArgumentException($"Unknown volume: '{p.Volume}'. Supported values: {SupportedVolumes}");

        var isOnClick = p.PlayMode.Equals("onclick", StringComparison.OrdinalIgnoreCase);

        if (shape is IAudioFrame audio)
        {
            audio.PlayMode = isOnClick ? AudioPlayModePreset.OnClick : AudioPlayModePreset.Auto;
            audio.Volume = volume;
            audio.PlayLoopMode = p.Loop;
        }
        else if (shape is IVideoFrame video)
        {
            video.PlayMode = isOnClick ? VideoPlayModePreset.OnClick : VideoPlayModePreset.Auto;
            video.Volume = volume;
            video.PlayLoopMode = p.Loop;
            video.RewindVideo = p.Rewind;
        }
        else
        {
            throw new ArgumentException($"Shape at index {p.ShapeIndex} is not an audio or video frame");
        }

        MarkModified(context);

        List<string> settings = [$"playMode={p.PlayMode}", $"volume={p.Volume}"];
        if (p.Loop) settings.Add("loop=true");
        if (p.Rewind && shape is IVideoFrame) settings.Add("rewind=true");

        return Success($"Playback settings updated ({string.Join(", ", settings)}).");
    }

    /// <summary>
    ///     Extracts playback parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted playback parameters.</returns>
    private static PlaybackParameters ExtractPlaybackParameters(OperationParameters parameters)
    {
        return new PlaybackParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetOptional<int?>("shapeIndex"),
            parameters.GetOptional("playMode", "auto"),
            parameters.GetOptional("loop", false),
            parameters.GetOptional("rewind", false),
            parameters.GetOptional("volume", "medium"));
    }

    /// <summary>
    ///     Record for holding playback parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="PlayMode">The play mode.</param>
    /// <param name="Loop">Whether to loop.</param>
    /// <param name="Rewind">Whether to rewind.</param>
    /// <param name="Volume">The volume level.</param>
    private record PlaybackParameters(
        int SlideIndex,
        int? ShapeIndex,
        string PlayMode,
        bool Loop,
        bool Rewind,
        string Volume);
}
