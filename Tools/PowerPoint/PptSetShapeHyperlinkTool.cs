using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptSetShapeHyperlinkTool : IAsposeTool
{
    public string Description => "Set a hyperlink on any shape (auto-shape, picture, etc.)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            slideIndex = new { type = "number", description = "Slide index (0-based)" },
            shapeIndex = new { type = "number", description = "Shape index (0-based)" },
            url = new { type = "string", description = "Hyperlink URL" },
            tooltip = new { type = "string", description = "Tooltip (optional)" }
        },
        required = new[] { "path", "slideIndex", "shapeIndex", "url" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var url = arguments?["url"]?.GetValue<string>() ?? throw new ArgumentException("url is required");
        var tooltip = arguments?["tooltip"]?.GetValue<string>();

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
        shape.HyperlinkClick = new Hyperlink(url) { Tooltip = tooltip };

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已設定超連結到 shape {shapeIndex}：{url}");
    }
}

