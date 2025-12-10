using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddHyperlinkTool : IAsposeTool
{
    public string Description => "Add a hyperlink text box to a slide";

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
                description = "Display text"
            },
            url = new
            {
                type = "string",
                description = "Hyperlink URL"
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
                description = "Width (optional, default: 300)"
            },
            height = new
            {
                type = "number",
                description = "Height (optional, default: 50)"
            }
        },
        required = new[] { "path", "slideIndex", "text", "url" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var url = arguments?["url"]?.GetValue<string>() ?? throw new ArgumentException("url is required");
        var x = arguments?["x"]?.GetValue<float?>() ?? 50;
        var y = arguments?["y"]?.GetValue<float?>() ?? 50;
        var width = arguments?["width"]?.GetValue<float?>() ?? 300;
        var height = arguments?["height"]?.GetValue<float?>() ?? 50;

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        var shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
        var textFrame = shape.TextFrame;
        textFrame.Text = text;

        var portion = textFrame.Paragraphs[0].Portions[0];
        portion.PortionFormat.HyperlinkClick = new Hyperlink(url);
        portion.PortionFormat.FontHeight = 14;
        portion.PortionFormat.FillFormat.FillType = FillType.Solid;
        portion.PortionFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Blue;

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已在投影片 {slideIndex} 新增超連結文字: {url}");
    }
}

