using System.Text.Json.Nodes;
using System.Text;
using System.Drawing;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint shape format (set, get)
/// Merges: PptSetShapeFormatTool, PptGetShapeFormatTool
/// </summary>
public class PptShapeFormatTool : IAsposeTool
{
    public string Description => "Manage PowerPoint shape format: set or get";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'set', 'get'",
                @enum = new[] { "set", "get" }
            },
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
                description = "X position (points, optional, for set)"
            },
            y = new
            {
                type = "number",
                description = "Y position (points, optional, for set)"
            },
            width = new
            {
                type = "number",
                description = "Width (points, optional, for set)"
            },
            height = new
            {
                type = "number",
                description = "Height (points, optional, for set)"
            },
            rotation = new
            {
                type = "number",
                description = "Rotation in degrees (optional, for set)"
            },
            fillColor = new
            {
                type = "string",
                description = "Fill color hex, e.g. #FFAA00 (optional, for set)"
            },
            lineColor = new
            {
                type = "string",
                description = "Line color hex (optional, for set)"
            }
        },
        required = new[] { "operation", "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");

        return operation.ToLower() switch
        {
            "set" => await SetShapeFormatAsync(arguments, path, slideIndex, shapeIndex),
            "get" => await GetShapeFormatAsync(arguments, path, slideIndex, shapeIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> SetShapeFormatAsync(JsonObject? arguments, string path, int slideIndex, int shapeIndex)
    {
        var x = arguments?["x"]?.GetValue<float?>();
        var y = arguments?["y"]?.GetValue<float?>();
        var width = arguments?["width"]?.GetValue<float?>();
        var height = arguments?["height"]?.GetValue<float?>();
        var rotation = arguments?["rotation"]?.GetValue<float?>();
        var fillColor = arguments?["fillColor"]?.GetValue<string>();
        var lineColor = arguments?["lineColor"]?.GetValue<string>();

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);

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

    private async Task<string> GetShapeFormatAsync(JsonObject? arguments, string path, int slideIndex, int shapeIndex)
    {
        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        var sb = new StringBuilder();

        sb.AppendLine($"Shape [{shapeIndex}] Format:");
        sb.AppendLine($"  Type: {shape.GetType().Name}");
        sb.AppendLine($"  Position: X={shape.X}, Y={shape.Y}");
        sb.AppendLine($"  Size: Width={shape.Width}, Height={shape.Height}");
        sb.AppendLine($"  Rotation: {shape.Rotation}°");

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

