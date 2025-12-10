using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAlignShapesTool : IAsposeTool
{
    public string Description => "Align multiple shapes (left/center/right/top/middle/bottom)";

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
            shapeIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Shape indices to align (0-based, at least 2)"
            },
            align = new
            {
                type = "string",
                description = "Alignment: left|center|right|top|middle|bottom"
            },
            alignToSlide = new
            {
                type = "boolean",
                description = "Align to slide instead of group (default: false)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndices", "align" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var alignStr = arguments?["align"]?.GetValue<string>() ?? throw new ArgumentException("align is required");
        var shapeIndices = arguments?["shapeIndices"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray()
                           ?? throw new ArgumentException("shapeIndices is required");
        var alignToSlide = arguments?["alignToSlide"]?.GetValue<bool?>() ?? false;

        if (shapeIndices.Length < 2) throw new ArgumentException("shapeIndices must contain at least 2 items");

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        foreach (var idx in shapeIndices)
        {
            if (idx < 0 || idx >= slide.Shapes.Count)
            {
                throw new ArgumentException($"shape index {idx} is out of range (0-{slide.Shapes.Count - 1})");
            }
        }

        var shapes = shapeIndices.Select(idx => slide.Shapes[idx]).ToArray();
        var refBox = alignToSlide
            ? new { X = 0f, Y = 0f, W = presentation.SlideSize.Size.Width, H = presentation.SlideSize.Size.Height }
            : new
            {
                X = shapes.Min(s => s.X),
                Y = shapes.Min(s => s.Y),
                W = shapes.Max(s => s.X + s.Width) - shapes.Min(s => s.X),
                H = shapes.Max(s => s.Y + s.Height) - shapes.Min(s => s.Y)
            };

        foreach (var s in shapes)
        {
            switch (alignStr.ToLower())
            {
                case "left":
                    s.X = refBox.X;
                    break;
                case "center":
                    s.X = refBox.X + (refBox.W - s.Width) / 2f;
                    break;
                case "right":
                    s.X = refBox.X + refBox.W - s.Width;
                    break;
                case "top":
                    s.Y = refBox.Y;
                    break;
                case "middle":
                    s.Y = refBox.Y + (refBox.H - s.Height) / 2f;
                    break;
                case "bottom":
                    s.Y = refBox.Y + refBox.H - s.Height;
                    break;
                default:
                    throw new ArgumentException("align must be one of: left, center, right, top, middle, bottom");
            }
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已對齊 {shapeIndices.Length} 個形狀：{alignStr}, alignToSlide={alignToSlide}");
    }
}

