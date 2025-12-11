using System.Text.Json.Nodes;
using System.Text;
using System.Drawing;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint backgrounds (set, get)
/// Merges: PptSetBackgroundTool, PptGetBackgroundTool
/// </summary>
public class PptBackgroundTool : IAsposeTool
{
    public string Description => "Manage PowerPoint backgrounds: set or get";

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
                description = "Slide index (0-based, optional, default: 0)"
            },
            color = new
            {
                type = "string",
                description = "Hex color like #FFAA00 (optional, for set)"
            },
            imagePath = new
            {
                type = "string",
                description = "Background image path (optional, for set)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        return operation.ToLower() switch
        {
            "set" => await SetBackgroundAsync(arguments, path),
            "get" => await GetBackgroundAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> SetBackgroundAsync(JsonObject? arguments, string path)
    {
        var slideIndex = arguments?["slideIndex"]?.GetValue<int?>() ?? 0;
        var colorHex = arguments?["color"]?.GetValue<string>();
        var imagePath = arguments?["imagePath"]?.GetValue<string>();

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var fillFormat = slide.Background.FillFormat;

        if (!string.IsNullOrWhiteSpace(imagePath))
        {
            var img = presentation.Images.AddImage(File.ReadAllBytes(imagePath));
            fillFormat.FillType = FillType.Picture;
            fillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
            fillFormat.PictureFillFormat.Picture.Image = img;
        }
        else if (!string.IsNullOrWhiteSpace(colorHex))
        {
            var color = ColorTranslator.FromHtml(colorHex);
            fillFormat.FillType = FillType.Solid;
            fillFormat.SolidFillColor.Color = color;
        }
        else
        {
            throw new ArgumentException("請至少提供 color 或 imagePath 其中之一");
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已更新投影片 {slideIndex} 背景: {path}");
    }

    private async Task<string> GetBackgroundAsync(JsonObject? arguments, string path)
    {
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required for get operation");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
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

