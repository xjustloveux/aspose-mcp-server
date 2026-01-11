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
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");
        var playModeStr = parameters.GetOptional("playMode", "auto");
        var loop = parameters.GetOptional("loop", false);
        var rewind = parameters.GetOptional("rewind", false);
        var volumeStr = parameters.GetOptional("volume", "medium");

        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for set_playback operation");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

        if (!VolumeMap.TryGetValue(volumeStr, out var volume))
            throw new ArgumentException($"Unknown volume: '{volumeStr}'. Supported values: {SupportedVolumes}");

        var isOnClick = playModeStr.Equals("onclick", StringComparison.OrdinalIgnoreCase);

        if (shape is IAudioFrame audio)
        {
            audio.PlayMode = isOnClick ? AudioPlayModePreset.OnClick : AudioPlayModePreset.Auto;
            audio.Volume = volume;
            audio.PlayLoopMode = loop;
        }
        else if (shape is IVideoFrame video)
        {
            video.PlayMode = isOnClick ? VideoPlayModePreset.OnClick : VideoPlayModePreset.Auto;
            video.Volume = volume;
            video.PlayLoopMode = loop;
            video.RewindVideo = rewind;
        }
        else
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not an audio or video frame");
        }

        MarkModified(context);

        List<string> settings = [$"playMode={playModeStr}", $"volume={volumeStr}"];
        if (loop) settings.Add("loop=true");
        if (rewind && shape is IVideoFrame) settings.Add("rewind=true");

        return Success($"Playback settings updated ({string.Join(", ", settings)}).");
    }
}
