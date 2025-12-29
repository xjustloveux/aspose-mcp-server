using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint media.
///     Supports: add_audio, delete_audio, add_video, delete_video, set_playback
/// </summary>
public class PptMediaTool : IAsposeTool
{
    private static readonly Dictionary<string, AudioVolumeMode> VolumeMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["mute"] = AudioVolumeMode.Mute,
        ["low"] = AudioVolumeMode.Low,
        ["medium"] = AudioVolumeMode.Medium,
        ["loud"] = AudioVolumeMode.Loud
    };

    private static readonly string SupportedVolumes = string.Join(", ", VolumeMap.Keys);

    public string Description =>
        @"Manage PowerPoint media. Supports 5 operations: add_audio, delete_audio, add_video, delete_video, set_playback.

Coordinate system: Origin is top-left corner of slide. Units are in Points (1 inch = 72 points).
Standard slide size: 720 x 540 points (10 x 7.5 inches).

Usage examples:
- Add audio: ppt_media(operation='add_audio', path='presentation.pptx', slideIndex=0, audioPath='audio.mp3', x=100, y=100)
- Delete audio: ppt_media(operation='delete_audio', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Add video: ppt_media(operation='add_video', path='presentation.pptx', slideIndex=0, videoPath='video.mp4', x=100, y=100)
- Delete video: ppt_media(operation='delete_video', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Set playback: ppt_media(operation='set_playback', path='presentation.pptx', slideIndex=0, shapeIndex=0, playMode='auto', loop=true)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add_audio': Embed audio into slide (required: path, slideIndex, audioPath)
- 'delete_audio': Delete audio frame (required: path, slideIndex, shapeIndex)
- 'add_video': Embed video into slide (required: path, slideIndex, videoPath)
- 'delete_video': Delete video frame (required: path, slideIndex, shapeIndex)
- 'set_playback': Set playback options for audio/video (required: path, slideIndex, shapeIndex)",
                @enum = new[] { "add_audio", "delete_audio", "add_video", "delete_video", "set_playback" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required for all operations)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index (0-based, required for delete/set_playback)"
            },
            audioPath = new
            {
                type = "string",
                description = "Audio file path to embed (required for add_audio)"
            },
            videoPath = new
            {
                type = "string",
                description = "Video file path to embed (required for add_video)"
            },
            x = new
            {
                type = "number",
                description = "X position in points from top-left corner (optional, default: 50)"
            },
            y = new
            {
                type = "number",
                description = "Y position in points from top-left corner (optional, default: 50)"
            },
            width = new
            {
                type = "number",
                description = "Width in points (optional, default: 80 for audio, 320 for video)"
            },
            height = new
            {
                type = "number",
                description = "Height in points (optional, default: 80 for audio, 240 for video)"
            },
            playMode = new
            {
                type = "string",
                description = "Playback mode: auto|onclick (optional, default: auto)"
            },
            loop = new
            {
                type = "boolean",
                description = "Loop playback (optional, default: false)"
            },
            rewind = new
            {
                type = "boolean",
                description = "Rewind video after play (optional, default: false)"
            },
            volume = new
            {
                type = "string",
                description = "Volume level: mute|low|medium|loud (optional, default: medium)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "add_audio" => await AddAudioAsync(path, outputPath, slideIndex, arguments),
            "delete_audio" => await DeleteAudioAsync(path, outputPath, slideIndex, arguments),
            "add_video" => await AddVideoAsync(path, outputPath, slideIndex, arguments),
            "delete_video" => await DeleteVideoAsync(path, outputPath, slideIndex, arguments),
            "set_playback" => await SetPlaybackAsync(path, outputPath, slideIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Embeds audio into a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing audioPath, optional x, y, width, height.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    /// <exception cref="FileNotFoundException">Thrown when audio file is not found.</exception>
    private Task<string> AddAudioAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var audioPath = ArgumentHelper.GetString(arguments, "audioPath");
            var x = ArgumentHelper.GetFloat(arguments, "x", 50);
            var y = ArgumentHelper.GetFloat(arguments, "y", 50);
            var width = ArgumentHelper.GetFloat(arguments, "width", 80);
            var height = ArgumentHelper.GetFloat(arguments, "height", 80);

            if (!File.Exists(audioPath))
                throw new FileNotFoundException($"Audio file not found: {audioPath}");

            using var presentation = new Presentation(path);
            PowerPointHelper.ValidateCollectionIndex(slideIndex, presentation.Slides.Count, "slide");

            var slide = presentation.Slides[slideIndex];
            using var audioStream = File.OpenRead(audioPath);
            slide.Shapes.AddAudioFrameEmbedded(x, y, width, height, audioStream);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Audio embedded into slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes audio from a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex or shapeIndex is out of range, or shape is not an audio
    ///     frame.
    /// </exception>
    private Task<string> DeleteAudioAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            if (shape is not IAudioFrame)
                throw new ArgumentException($"Shape at index {shapeIndex} is not an audio frame");

            slide.Shapes.Remove(shape);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Audio deleted from slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Embeds video into a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing videoPath, optional x, y, width, height.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    /// <exception cref="FileNotFoundException">Thrown when video file is not found.</exception>
    private Task<string> AddVideoAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var videoPath = ArgumentHelper.GetString(arguments, "videoPath");
            var x = ArgumentHelper.GetFloat(arguments, "x", 50);
            var y = ArgumentHelper.GetFloat(arguments, "y", 50);
            var width = ArgumentHelper.GetFloat(arguments, "width", 320);
            var height = ArgumentHelper.GetFloat(arguments, "height", 240);

            if (!File.Exists(videoPath))
                throw new FileNotFoundException($"Video file not found: {videoPath}");

            using var presentation = new Presentation(path);
            PowerPointHelper.ValidateCollectionIndex(slideIndex, presentation.Slides.Count, "slide");

            var slide = presentation.Slides[slideIndex];
            using var videoStream = File.OpenRead(videoPath);
            var video = presentation.Videos.AddVideo(videoStream, LoadingStreamBehavior.KeepLocked);
            slide.Shapes.AddVideoFrame(x, y, width, height, video);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Video embedded into slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes video from a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex or shapeIndex is out of range, or shape is not a video
    ///     frame.
    /// </exception>
    private Task<string> DeleteVideoAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            if (shape is not IVideoFrame)
                throw new ArgumentException($"Shape at index {shapeIndex} is not a video frame");

            slide.Shapes.Remove(shape);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Video deleted from slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets media playback options for audio or video.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, optional playMode, loop, rewind, volume.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex or shapeIndex is out of range, shape is not a media frame,
    ///     or volume is invalid.
    /// </exception>
    private Task<string> SetPlaybackAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var playModeStr = ArgumentHelper.GetString(arguments, "playMode", "auto");
            var loop = ArgumentHelper.GetBool(arguments, "loop", false);
            var rewind = ArgumentHelper.GetBool(arguments, "rewind", false);
            var volumeStr = ArgumentHelper.GetString(arguments, "volume", "medium");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

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

            presentation.Save(outputPath, SaveFormat.Pptx);

            var settings = new List<string> { $"playMode={playModeStr}", $"volume={volumeStr}" };
            if (loop) settings.Add("loop=true");
            if (rewind && shape is IVideoFrame) settings.Add("rewind=true");

            return $"Playback settings updated ({string.Join(", ", settings)}). Output: {outputPath}";
        });
    }
}