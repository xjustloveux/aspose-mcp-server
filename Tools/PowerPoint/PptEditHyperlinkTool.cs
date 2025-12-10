using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptEditHyperlinkTool : IAsposeTool
{
    public string Description => "Edit hyperlink URL or target on a shape in PowerPoint";

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
            shapeIndex = new
            {
                type = "number",
                description = "Shape index (0-based)"
            },
            url = new
            {
                type = "string",
                description = "New hyperlink URL (optional)"
            },
            slideTargetIndex = new
            {
                type = "number",
                description = "Target slide index for internal link (optional)"
            },
            removeHyperlink = new
            {
                type = "boolean",
                description = "Remove hyperlink (optional)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var url = arguments?["url"]?.GetValue<string>();
        var slideTargetIndex = arguments?["slideTargetIndex"]?.GetValue<int?>();
        var removeHyperlink = arguments?["removeHyperlink"]?.GetValue<bool?>() ?? false;

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

        var shape = slide.Shapes[shapeIndex];

        if (removeHyperlink)
        {
            if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
            {
                foreach (var paragraph in autoShape.TextFrame.Paragraphs)
                {
                    foreach (var portion in paragraph.Portions)
                    {
                        portion.PortionFormat.HyperlinkClick = null;
                    }
                }
            }
            shape.HyperlinkClick = null;
        }
        else if (!string.IsNullOrEmpty(url))
        {
            shape.HyperlinkClick = new Hyperlink(url);
        }
        else if (slideTargetIndex.HasValue)
        {
            if (slideTargetIndex.Value < 0 || slideTargetIndex.Value >= presentation.Slides.Count)
            {
                throw new ArgumentException($"slideTargetIndex must be between 0 and {presentation.Slides.Count - 1}");
            }
            shape.HyperlinkClick = new Hyperlink(presentation.Slides[slideTargetIndex.Value]);
        }
        else
        {
            throw new ArgumentException("Either url, slideTargetIndex, or removeHyperlink must be provided");
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Hyperlink updated on slide {slideIndex}, shape {shapeIndex}");
    }
}

