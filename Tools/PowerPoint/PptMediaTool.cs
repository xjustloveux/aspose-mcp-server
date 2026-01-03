using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint media.
///     Supports: add_audio, delete_audio, add_video, delete_video, set_playback
/// </summary>
[McpServerToolType]
public class PptMediaTool
{
    /// <summary>
    ///     Mapping of volume level string names to AudioVolumeMode enum values.
    /// </summary>
    private static readonly Dictionary<string, AudioVolumeMode> VolumeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["mute"] = AudioVolumeMode.Mute,
        ["low"] = AudioVolumeMode.Low,
        ["medium"] = AudioVolumeMode.Medium,
        ["loud"] = AudioVolumeMode.Loud
    };

    /// <summary>
    ///     Comma-separated list of supported volume level names for error messages.
    /// </summary>
    private static readonly string SupportedVolumes = string.Join(", ", VolumeMap.Keys);

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptMediaTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PptMediaTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_media")]
    [Description(
        @"Manage PowerPoint media. Supports 5 operations: add_audio, delete_audio, add_video, delete_video, set_playback.

Coordinate system: Origin is top-left corner of slide. Units are in Points (1 inch = 72 points).
Standard slide size: 720 x 540 points (10 x 7.5 inches).

Usage examples:
- Add audio: ppt_media(operation='add_audio', path='presentation.pptx', slideIndex=0, audioPath='audio.mp3', x=100, y=100)
- Delete audio: ppt_media(operation='delete_audio', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Add video: ppt_media(operation='add_video', path='presentation.pptx', slideIndex=0, videoPath='video.mp4', x=100, y=100)
- Delete video: ppt_media(operation='delete_video', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Set playback: ppt_media(operation='set_playback', path='presentation.pptx', slideIndex=0, shapeIndex=0, playMode='auto', loop=true)")]
    public string Execute(
        [Description("Operation: add_audio, delete_audio, add_video, delete_video, set_playback")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for all operations)")]
        int slideIndex = 0,
        [Description("Shape index (0-based, required for delete/set_playback)")]
        int? shapeIndex = null,
        [Description("Audio file path to embed (required for add_audio)")]
        string? audioPath = null,
        [Description("Video file path to embed (required for add_video)")]
        string? videoPath = null,
        [Description("X position in points from top-left corner (optional, default: 50)")]
        float x = 50,
        [Description("Y position in points from top-left corner (optional, default: 50)")]
        float y = 50,
        [Description("Width in points (optional, default: 80 for audio, 320 for video)")]
        float? width = null,
        [Description("Height in points (optional, default: 80 for audio, 240 for video)")]
        float? height = null,
        [Description("Playback mode: auto|onclick (optional, default: auto)")]
        string playMode = "auto",
        [Description("Loop playback (optional, default: false)")]
        bool loop = false,
        [Description("Rewind video after play (optional, default: false)")]
        bool rewind = false,
        [Description("Volume level: mute|low|medium|loud (optional, default: medium)")]
        string volume = "medium")
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add_audio" => AddAudio(ctx, outputPath, slideIndex, audioPath, x, y, width ?? 80, height ?? 80),
            "delete_audio" => DeleteAudio(ctx, outputPath, slideIndex, shapeIndex),
            "add_video" => AddVideo(ctx, outputPath, slideIndex, videoPath, x, y, width ?? 320, height ?? 240),
            "delete_video" => DeleteVideo(ctx, outputPath, slideIndex, shapeIndex),
            "set_playback" => SetPlayback(ctx, outputPath, slideIndex, shapeIndex, playMode, loop, rewind, volume),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Embeds audio into a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="audioPath">The audio file path to embed.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when audioPath is not provided.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the audio file is not found.</exception>
    private static string AddAudio(DocumentContext<Presentation> ctx, string? outputPath,
        int slideIndex, string? audioPath, float x, float y, float width, float height)
    {
        if (string.IsNullOrEmpty(audioPath))
            throw new ArgumentException("audioPath is required for add_audio operation");
        if (!File.Exists(audioPath))
            throw new FileNotFoundException($"Audio file not found: {audioPath}");

        var presentation = ctx.Document;
        PowerPointHelper.ValidateCollectionIndex(slideIndex, presentation.Slides.Count, "slide");

        var slide = presentation.Slides[slideIndex];
        using var audioStream = File.OpenRead(audioPath);
        slide.Shapes.AddAudioFrameEmbedded(x, y, width, height, audioStream);

        ctx.Save(outputPath);

        var result = $"Audio embedded into slide {slideIndex}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Deletes audio from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided or the shape is not an audio frame.</exception>
    private static string DeleteAudio(DocumentContext<Presentation> ctx, string? outputPath,
        int slideIndex, int? shapeIndex)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete_audio operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

        if (shape is not IAudioFrame)
            throw new ArgumentException($"Shape at index {shapeIndex} is not an audio frame");

        slide.Shapes.Remove(shape);

        ctx.Save(outputPath);

        var result = $"Audio deleted from slide {slideIndex}, shape {shapeIndex}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Embeds video into a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="videoPath">The video file path to embed.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when videoPath is not provided.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the video file is not found.</exception>
    private static string AddVideo(DocumentContext<Presentation> ctx, string? outputPath,
        int slideIndex, string? videoPath, float x, float y, float width, float height)
    {
        if (string.IsNullOrEmpty(videoPath))
            throw new ArgumentException("videoPath is required for add_video operation");
        if (!File.Exists(videoPath))
            throw new FileNotFoundException($"Video file not found: {videoPath}");

        var presentation = ctx.Document;
        PowerPointHelper.ValidateCollectionIndex(slideIndex, presentation.Slides.Count, "slide");

        var slide = presentation.Slides[slideIndex];
        using var videoStream = File.OpenRead(videoPath);
        var video = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.KeepLocked);
        slide.Shapes.AddVideoFrame(x, y, width, height, video);

        ctx.Save(outputPath);

        var result = $"Video embedded into slide {slideIndex}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Deletes video from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided or the shape is not a video frame.</exception>
    private static string DeleteVideo(DocumentContext<Presentation> ctx, string? outputPath,
        int slideIndex, int? shapeIndex)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete_video operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

        if (shape is not IVideoFrame)
            throw new ArgumentException($"Shape at index {shapeIndex} is not a video frame");

        slide.Shapes.Remove(shape);

        ctx.Save(outputPath);

        var result = $"Video deleted from slide {slideIndex}, shape {shapeIndex}. ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets media playback options for audio or video.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="playModeStr">The playback mode (auto or onclick).</param>
    /// <param name="loop">Whether to loop playback.</param>
    /// <param name="rewind">Whether to rewind video after play.</param>
    /// <param name="volumeStr">The volume level string.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when shapeIndex is not provided, volume is invalid, or shape is not a media
    ///     frame.
    /// </exception>
    private static string SetPlayback(DocumentContext<Presentation> ctx, string? outputPath,
        int slideIndex, int? shapeIndex, string playModeStr, bool loop, bool rewind, string volumeStr)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for set_playback operation");

        var presentation = ctx.Document;
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

        ctx.Save(outputPath);

        List<string> settings = [$"playMode={playModeStr}", $"volume={volumeStr}"];
        if (loop) settings.Add("loop=true");
        if (rewind && shape is IVideoFrame) settings.Add("rewind=true");

        var result = $"Playback settings updated ({string.Join(", ", settings)}). ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }
}