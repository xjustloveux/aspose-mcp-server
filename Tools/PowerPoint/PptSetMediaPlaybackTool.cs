using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptSetMediaPlaybackTool : IAsposeTool
{
    public string Description => "Set playback options for audio/video shapes (auto/on-click, loop, volume)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            slideIndex = new { type = "number", description = "Slide index (0-based)" },
            shapeIndex = new { type = "number", description = "Audio/Video shape index (0-based)" },
            playMode = new { type = "string", description = "auto|onclick (default: auto)" },
            loop = new { type = "boolean", description = "Loop playback (default: false)" },
            rewind = new { type = "boolean", description = "Rewind after play (video) (default: false)" },
            volume = new { type = "string", description = "mute|low|medium|loud (default: medium)" }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
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

