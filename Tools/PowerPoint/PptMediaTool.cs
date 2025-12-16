using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint media (add audio/video, delete, set playback)
/// Merges: PptAddAudioTool, PptDeleteAudioTool, PptAddVideoTool, PptDeleteVideoTool, PptSetMediaPlaybackTool
/// </summary>
public class PptMediaTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint media. Supports 5 operations: add_audio, delete_audio, add_video, delete_video, set_playback.

Usage examples:
- Add audio: ppt_media(operation='add_audio', path='presentation.pptx', slideIndex=0, audioPath='audio.mp3', x=100, y=100)
- Delete audio: ppt_media(operation='delete_audio', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Add video: ppt_media(operation='add_video', path='presentation.pptx', slideIndex=0, videoPath='video.mp4', x=100, y=100)
- Delete video: ppt_media(operation='delete_video', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Set playback: ppt_media(operation='set_playback', path='presentation.pptx', slideIndex=0, shapeIndex=0, playOnClick=true)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add_audio': Add audio to slide (required params: path, slideIndex, audioPath)
- 'delete_audio': Delete audio (required params: path, slideIndex, shapeIndex)
- 'add_video': Add video to slide (required params: path, slideIndex, videoPath)
- 'delete_video': Delete video (required params: path, slideIndex, shapeIndex)
- 'set_playback': Set playback options (required params: path, slideIndex, shapeIndex)",
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
                description = "Slide index (0-based)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index (0-based, required for delete_audio/delete_video/set_playback)"
            },
            audioPath = new
            {
                type = "string",
                description = "Audio file path (required for add_audio)"
            },
            videoPath = new
            {
                type = "string",
                description = "Video file path (required for add_video)"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, for add_audio/add_video)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, for add_audio/add_video)"
            },
            width = new
            {
                type = "number",
                description = "Width (optional, for add_audio/add_video)"
            },
            height = new
            {
                type = "number",
                description = "Height (optional, for add_audio/add_video)"
            },
            playMode = new
            {
                type = "string",
                description = "auto|onclick (optional, for set_playback, default: auto)"
            },
            loop = new
            {
                type = "boolean",
                description = "Loop playback (optional, for set_playback, default: false)"
            },
            rewind = new
            {
                type = "boolean",
                description = "Rewind after play (video) (optional, for set_playback, default: false)"
            },
            volume = new
            {
                type = "string",
                description = "mute|low|medium|loud (optional, for set_playback, default: medium)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete/set_playback operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "add_audio" => await AddAudioAsync(arguments, path, slideIndex),
            "delete_audio" => await DeleteAudioAsync(arguments, path, slideIndex),
            "add_video" => await AddVideoAsync(arguments, path, slideIndex),
            "delete_video" => await DeleteVideoAsync(arguments, path, slideIndex),
            "set_playback" => await SetPlaybackAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds audio to a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing audioPath, optional x, y, width, height, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> AddAudioAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var audioPath = ArgumentHelper.GetString(arguments, "audioPath");
        var x = ArgumentHelper.GetFloat(arguments, "x", 50);
        var y = ArgumentHelper.GetFloat(arguments, "y", 50);
        var width = ArgumentHelper.GetFloat(arguments, "width", 80);
        var height = ArgumentHelper.GetFloat(arguments, "height", 80);

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        using var audioStream = File.OpenRead(audioPath);
        slide.Shapes.AddAudioFrameEmbedded(x, y, width, height, audioStream);

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Audio inserted into slide {slideIndex}: {audioPath}");
    }

    /// <summary>
    /// Deletes audio from a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing audioIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteAudioAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        if (shape is not IAudioFrame)
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not an audio frame");
        }

        slide.Shapes.Remove(shape);

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Audio deleted from slide {slideIndex}, shape {shapeIndex}");
    }

    /// <summary>
    /// Adds video to a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing videoPath, optional x, y, width, height, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> AddVideoAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var videoPath = ArgumentHelper.GetString(arguments, "videoPath");
        var x = ArgumentHelper.GetFloat(arguments, "x", 50);
        var y = ArgumentHelper.GetFloat(arguments, "y", 50);
        var width = ArgumentHelper.GetFloat(arguments, "width", 320);
        var height = ArgumentHelper.GetFloat(arguments, "height", 240);

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var frame = slide.Shapes.AddVideoFrame(x, y, width, height, videoPath);

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Video inserted into slide {slideIndex}: {videoPath}");
    }

    /// <summary>
    /// Deletes video from a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing videoIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteVideoAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        if (shape is not IVideoFrame)
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not a video frame");
        }

        slide.Shapes.Remove(shape);

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Video deleted from slide {slideIndex}, shape {shapeIndex}");
    }

    /// <summary>
    /// Sets media playback options
    /// </summary>
    /// <param name="arguments">JSON arguments containing mediaIndex, optional playMode, loop, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> SetPlaybackAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
        var playModeStr = ArgumentHelper.GetString(arguments, "playMode", "auto");
        var loop = ArgumentHelper.GetBool(arguments, "loop", false);
        var rewind = ArgumentHelper.GetBool(arguments, "rewind", false);
        var volumeStr = ArgumentHelper.GetString(arguments, "volume", "medium");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");
        }

        var playModeAudio = playModeStr.ToLower() == "onclick" ? AudioPlayModePreset.OnClick : AudioPlayModePreset.Auto;
        var playModeVideo = playModeStr.ToLower() == "onclick" ? VideoPlayModePreset.OnClick : VideoPlayModePreset.Auto;
        var volume = volumeStr.ToLower() switch
        {
            "mute" => AudioVolumeMode.Mute,
            "low" => AudioVolumeMode.Low,
            "loud" => AudioVolumeMode.Loud,
            _ => AudioVolumeMode.Medium
        };

        var shape = slide.Shapes[shapeIndex];
        if (shape is IAudioFrame audio)
        {
            audio.PlayMode = playModeAudio;
            audio.Volume = volume;
            audio.PlayLoopMode = loop;
        }
        else if (shape is IVideoFrame video)
        {
            video.PlayMode = playModeVideo;
            video.Volume = volume;
            video.PlayLoopMode = loop;
            video.RewindVideo = rewind;
        }
        else
        {
            throw new ArgumentException("The specified shape is not audio or video");
        }

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Media playback settings updated: slide {slideIndex}, shape {shapeIndex}");
    }
}

