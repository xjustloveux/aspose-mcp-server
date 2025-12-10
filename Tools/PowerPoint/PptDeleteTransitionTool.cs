using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;

namespace AsposeMcpServer.Tools;

public class PptDeleteTransitionTool : IAsposeTool
{
    public string Description => "Remove transition from a slide";

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
        slide.SlideShowTransition.Type = TransitionType.Fade;
        slide.SlideShowTransition.AdvanceOnClick = false;
        slide.SlideShowTransition.AdvanceAfterTime = 0;

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Transition removed from slide {slideIndex}: {path}");
    }
}

