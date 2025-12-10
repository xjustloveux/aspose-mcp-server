using System.Text.Json.Nodes;
using System.Text;
using System.Drawing;
using Aspose.Slides;

namespace AsposeMcpServer.Tools;

public class PptGetShapeFormatTool : IAsposeTool
{
    public string Description => "Get detailed format information for a shape on a PowerPoint slide";

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
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");

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
        var sb = new StringBuilder();

        sb.AppendLine($"Shape [{shapeIndex}] Format:");
        sb.AppendLine($"  Type: {shape.GetType().Name}");
        sb.AppendLine($"  Position: X={shape.X}, Y={shape.Y}");
        sb.AppendLine($"  Size: Width={shape.Width}, Height={shape.Height}");
        sb.AppendLine($"  Rotation: {shape.Rotation}°");
        // Note: Flip properties may not be directly available on IShape

        // Fill format
        sb.AppendLine($"  Fill Type: {shape.FillFormat.FillType}");
        if (shape.FillFormat.FillType == FillType.Solid)
        {
            var color = shape.FillFormat.SolidFillColor.Color;
            sb.AppendLine($"  Fill Color: RGB({color.R}, {color.G}, {color.B}), Hex: #{color.R:X2}{color.G:X2}{color.B:X2}");
        }

        // Line format
        sb.AppendLine($"  Line Width: {shape.LineFormat.Width}");
        if (shape.LineFormat.FillFormat.FillType == FillType.Solid)
        {
            var color = shape.LineFormat.FillFormat.SolidFillColor.Color;
            sb.AppendLine($"  Line Color: RGB({color.R}, {color.G}, {color.B}), Hex: #{color.R:X2}{color.G:X2}{color.B:X2}");
        }

        // Text format (if applicable)
        if (shape is IAutoShape autoShape && autoShape.TextFrame != null)
        {
            sb.AppendLine($"  Text: {autoShape.TextFrame.Text}");
            if (autoShape.TextFrame.Paragraphs.Count > 0)
            {
                var firstPara = autoShape.TextFrame.Paragraphs[0];
                if (firstPara.Portions.Count > 0)
                {
                    var portion = firstPara.Portions[0];
                    sb.AppendLine($"  Font: {portion.PortionFormat.LatinFont?.FontName ?? "(default)"}");
                    sb.AppendLine($"  Font Size: {portion.PortionFormat.FontHeight}");
                    sb.AppendLine($"  Bold: {portion.PortionFormat.FontBold}, Italic: {portion.PortionFormat.FontItalic}");
                }
            }
        }

        // Hyperlink
        if (shape.HyperlinkClick != null)
        {
            var url = shape.HyperlinkClick.ExternalUrl ?? (shape.HyperlinkClick.TargetSlide != null ? $"Slide {presentation.Slides.IndexOf(shape.HyperlinkClick.TargetSlide)}" : "Internal link");
            sb.AppendLine($"  Hyperlink: {url}");
        }

        return await Task.FromResult(sb.ToString());
    }
}

