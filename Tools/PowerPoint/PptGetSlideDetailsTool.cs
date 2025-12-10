using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGetSlideDetailsTool : IAsposeTool
{
    public string Description => "Get detailed information about a slide (transition, animations, background, etc.)";

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
        var sb = new StringBuilder();

        sb.AppendLine($"=== Slide {slideIndex} Details ===");
        sb.AppendLine($"Hidden: {slide.Hidden}");
        sb.AppendLine($"Layout: {slide.LayoutSlide?.Name ?? "None"}");
        sb.AppendLine($"Shapes Count: {slide.Shapes.Count}");

        // Transition
        var transition = slide.SlideShowTransition;
        if (transition != null)
        {
            sb.AppendLine($"\nTransition:");
            sb.AppendLine($"  Type: {transition.Type}");
            sb.AppendLine($"  Speed: {transition.Speed}");
            sb.AppendLine($"  AdvanceOnClick: {transition.AdvanceOnClick}");
            sb.AppendLine($"  AdvanceAfterTime: {transition.AdvanceAfterTime}ms");
        }

        // Animations
        var animations = slide.Timeline.MainSequence;
        sb.AppendLine($"\nAnimations: {animations.Count}");
        for (int i = 0; i < animations.Count; i++)
        {
            var anim = animations[i];
            sb.AppendLine($"  [{i}] Type: {anim.Type}, Shape: {anim.TargetShape?.GetType().Name}");
        }

        // Background
        var background = slide.Background;
        if (background != null)
        {
            sb.AppendLine($"\nBackground:");
            sb.AppendLine($"  FillType: {background.FillFormat.FillType}");
        }

        // Notes
        var notesSlide = slide.NotesSlideManager.NotesSlide;
        if (notesSlide != null)
        {
            var notesText = notesSlide.NotesTextFrame?.Text;
            sb.AppendLine($"\nNotes: {(string.IsNullOrWhiteSpace(notesText) ? "None" : notesText)}");
        }

        return await Task.FromResult(sb.ToString());
    }
}

