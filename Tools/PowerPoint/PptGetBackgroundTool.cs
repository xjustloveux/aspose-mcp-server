using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGetBackgroundTool : IAsposeTool
{
    public string Description => "Get background information for a slide";

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
        var background = slide.Background;
        var sb = new StringBuilder();

        sb.AppendLine($"=== Slide {slideIndex} Background ===");
        if (background != null)
        {
            sb.AppendLine($"FillType: {background.FillFormat.FillType}");
            if (background.FillFormat.FillType == FillType.Solid)
            {
                sb.AppendLine($"Color: {background.FillFormat.SolidFillColor}");
            }
            else if (background.FillFormat.FillType == FillType.Picture)
            {
                sb.AppendLine("Picture fill");
            }
        }
        else
        {
            sb.AppendLine("No background set");
        }

        return await Task.FromResult(sb.ToString());
    }
}

