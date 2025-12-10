using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddTextTool : IAsposeTool
{
    public string Description => "Add text to a PowerPoint slide";

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
            text = new
            {
                type = "string",
                description = "Text to add"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, default: 50)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, default: 50)"
            },
            width = new
            {
                type = "number",
                description = "Text box width (optional, default: 400)"
            },
            height = new
            {
                type = "number",
                description = "Text box height (optional, default: 100)"
            }
        },
        required = new[] { "path", "slideIndex", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var x = arguments?["x"]?.GetValue<float>() ?? 50;
        var y = arguments?["y"]?.GetValue<float>() ?? 50;
        var width = arguments?["width"]?.GetValue<float>() ?? 400;
        var height = arguments?["height"]?.GetValue<float>() ?? 100;

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
        textBox.TextFrame.Text = text;

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Text added to slide {slideIndex}: {path}");
    }
}

