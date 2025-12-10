using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGetTransitionTool : IAsposeTool
{
    public string Description => "Get transition information for a slide";

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
            }
        },
        required = new[] { "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var transition = slide.SlideShowTransition;
        var sb = new StringBuilder();

        sb.AppendLine($"=== Slide {slideIndex} Transition ===");
        if (transition != null)
        {
            sb.AppendLine($"Type: {transition.Type}");
            sb.AppendLine($"Speed: {transition.Speed}");
            sb.AppendLine($"AdvanceOnClick: {transition.AdvanceOnClick}");
            sb.AppendLine($"AdvanceAfterTime: {transition.AdvanceAfterTime}ms");
            sb.AppendLine($"SoundMode: {transition.SoundMode}");
            if (transition.Sound != null)
            {
                sb.AppendLine($"Sound: {transition.Sound}");
            }
        }
        else
        {
            sb.AppendLine("No transition set");
        }

        return await Task.FromResult(sb.ToString());
    }
}

