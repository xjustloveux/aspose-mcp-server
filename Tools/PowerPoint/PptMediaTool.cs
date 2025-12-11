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
    public string Description => "Manage PowerPoint media: add audio/video, delete, or set playback options";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add_audio', 'delete_audio', 'add_video', 'delete_video', 'set_playback'",
                @enum = new[] { "add_audio", "delete_audio", "add_video", "delete_video", "set_playback" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path"
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
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");

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

    private async Task<string> AddAudioAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var audioPath = arguments?["audioPath"]?.GetValue<string>() ?? throw new ArgumentException("audioPath is required for add_audio operation");
        var x = arguments?["x"]?.GetValue<float?>() ?? 50;
        var y = arguments?["y"]?.GetValue<float?>() ?? 50;
        var width = arguments?["width"]?.GetValue<float?>() ?? 80;
        var height = arguments?["height"]?.GetValue<float?>() ?? 80;

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        using var audioStream = File.OpenRead(audioPath);
        slide.Shapes.AddAudioFrameEmbedded(x, y, width, height, audioStream);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已在投影片 {slideIndex} 插入音訊: {audioPath}");
    }

    private async Task<string> DeleteAudioAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for delete_audio operation");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        if (shape is not IAudioFrame)
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not an audio frame");
        }

        slide.Shapes.Remove(shape);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Audio deleted from slide {slideIndex}, shape {shapeIndex}");
    }

    private async Task<string> AddVideoAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var videoPath = arguments?["videoPath"]?.GetValue<string>() ?? throw new ArgumentException("videoPath is required for add_video operation");
        var x = arguments?["x"]?.GetValue<float?>() ?? 50;
        var y = arguments?["y"]?.GetValue<float?>() ?? 50;
        var width = arguments?["width"]?.GetValue<float?>() ?? 320;
        var height = arguments?["height"]?.GetValue<float?>() ?? 240;

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var frame = slide.Shapes.AddVideoFrame(x, y, width, height, videoPath);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已在投影片 {slideIndex} 插入影片: {videoPath}");
    }

    private async Task<string> DeleteVideoAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for delete_video operation");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        if (shape is not IVideoFrame)
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not a video frame");
        }

        slide.Shapes.Remove(shape);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Video deleted from slide {slideIndex}, shape {shapeIndex}");
    }

    private async Task<string> SetPlaybackAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for set_playback operation");
        var playModeStr = arguments?["playMode"]?.GetValue<string>() ?? "auto";
        var loop = arguments?["loop"]?.GetValue<bool?>() ?? false;
        var rewind = arguments?["rewind"]?.GetValue<bool?>() ?? false;
        var volumeStr = arguments?["volume"]?.GetValue<string>() ?? "medium";

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
            throw new ArgumentException("指定的 shape 不是音訊或影片");
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已更新媒體播放設定：slide {slideIndex}, shape {shapeIndex}");
    }
}

