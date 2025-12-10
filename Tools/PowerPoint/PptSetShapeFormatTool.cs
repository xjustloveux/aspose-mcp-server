using System.Text.Json.Nodes;
using System.Drawing;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptSetShapeFormatTool : IAsposeTool
{
    public string Description => "Set shape position/size/rotation and fill/line colors";

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
            x = new
            {
                type = "number",
                description = "X position (points, optional)"
            },
            y = new
            {
                type = "number",
                description = "Y position (points, optional)"
            },
            width = new
            {
                type = "number",
                description = "Width (points, optional)"
            },
            height = new
            {
                type = "number",
                description = "Height (points, optional)"
            },
            rotation = new
            {
                type = "number",
                description = "Rotation in degrees (optional)"
            },
            fillColor = new
            {
                type = "string",
                description = "Fill color hex, e.g. #FFAA00 (optional)"
            },
            lineColor = new
            {
                type = "string",
                description = "Line color hex (optional)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var x = arguments?["x"]?.GetValue<float?>();
        var y = arguments?["y"]?.GetValue<float?>();
        var width = arguments?["width"]?.GetValue<float?>();
        var height = arguments?["height"]?.GetValue<float?>();
        var rotation = arguments?["rotation"]?.GetValue<float?>();
        var fillColor = arguments?["fillColor"]?.GetValue<string>();
        var lineColor = arguments?["lineColor"]?.GetValue<string>();

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

        if (x.HasValue) shape.X = x.Value;
        if (y.HasValue) shape.Y = y.Value;
        if (width.HasValue) shape.Width = width.Value;
        if (height.HasValue) shape.Height = height.Value;
        if (rotation.HasValue) shape.Rotation = rotation.Value;

        if (!string.IsNullOrWhiteSpace(fillColor))
        {
            var color = ColorTranslator.FromHtml(fillColor);
            shape.FillFormat.FillType = FillType.Solid;
            shape.FillFormat.SolidFillColor.Color = color;
        }

        if (!string.IsNullOrWhiteSpace(lineColor))
        {
            var color = ColorTranslator.FromHtml(lineColor);
            shape.LineFormat.FillFormat.FillType = FillType.Solid;
            shape.LineFormat.FillFormat.SolidFillColor.Color = color;
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已更新形狀格式：slide {slideIndex}, shape {shapeIndex}");
    }
}

