using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGetHyperlinksTool : IAsposeTool
{
    public string Description => "Get all hyperlinks in a presentation or on a specific slide";

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
                description = "Slide index (0-based, optional, if not provided gets all slides)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        var sb = new StringBuilder();

        if (slideIndex.HasValue)
        {
            if (slideIndex.Value < 0 || slideIndex.Value >= presentation.Slides.Count)
            {
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
            }
            var slide = presentation.Slides[slideIndex.Value];
            sb.AppendLine($"=== Slide {slideIndex.Value} Hyperlinks ===");
            GetHyperlinksFromSlide(presentation, slide, sb);
        }
        else
        {
            sb.AppendLine("=== All Hyperlinks ===");
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                var slide = presentation.Slides[i];
                var hyperlinks = GetHyperlinksFromSlide(presentation, slide, null);
                if (hyperlinks > 0)
                {
                    sb.AppendLine($"\nSlide {i}: {hyperlinks} hyperlink(s)");
                    GetHyperlinksFromSlide(presentation, slide, sb);
                }
            }
        }

        return await Task.FromResult(sb.ToString());
    }

    private int GetHyperlinksFromSlide(IPresentation presentation, ISlide slide, StringBuilder? sb)
    {
        int count = 0;
        foreach (var shape in slide.Shapes)
        {
            if (shape is IAutoShape autoShape)
            {
                if (autoShape.HyperlinkClick != null)
                {
                    count++;
                    if (sb != null)
                    {
                        var url = autoShape.HyperlinkClick.ExternalUrl ?? (autoShape.HyperlinkClick.TargetSlide != null ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkClick.TargetSlide)}" : "Internal link");
                        sb.AppendLine($"  Shape [{slide.Shapes.IndexOf(shape)}]: {url}");
                    }
                }
                if (autoShape.HyperlinkMouseOver != null)
                {
                    count++;
                    if (sb != null)
                    {
                        var url = autoShape.HyperlinkMouseOver.ExternalUrl ?? (autoShape.HyperlinkMouseOver.TargetSlide != null ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkMouseOver.TargetSlide)}" : "Internal link");
                        sb.AppendLine($"  Shape [{slide.Shapes.IndexOf(shape)}] (mouseover): {url}");
                    }
                }
            }
        }
        return count;
    }
}
