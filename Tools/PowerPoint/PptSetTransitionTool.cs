using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;

namespace AsposeMcpServer.Tools;

public class PptSetTransitionTool : IAsposeTool
{
    public string Description => "Set slide transition type and duration";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
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
            transitionType = new
            {
                type = "string",
                description = "Transition type (Fade, Push, Wipe, Split, RandomBars, etc.)"
            },
            durationSeconds = new
            {
                type = "number",
                description = "Transition duration in seconds (optional, default 1.0)"
            }
        },
        required = new[] { "path", "slideIndex", "transitionType" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var transitionTypeStr = arguments?["transitionType"]?.GetValue<string>() ?? throw new ArgumentException("transitionType is required");
        var duration = arguments?["durationSeconds"]?.GetValue<double?>() ?? 1.0;

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var transition = slide.SlideShowTransition;
        transition.Type = transitionTypeStr.ToLower() switch
        {
            "push" => TransitionType.Push,
            "wipe" => TransitionType.Wipe,
            "split" => TransitionType.Split,
            "randombars" => TransitionType.Random,
            "circle" => TransitionType.Circle,
            "plus" => TransitionType.Plus,
            "diamond" => TransitionType.Diamond,
            "fade" => TransitionType.Fade,
            _ => TransitionType.Fade
        };
        transition.AdvanceAfterTime = (uint)(duration * 1000);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已設定投影片 {slideIndex} 轉場：{transition.Type}，時間 {duration:0.##}s");
    }
}

