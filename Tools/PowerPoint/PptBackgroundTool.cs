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
    public string Description => @"Manage PowerPoint backgrounds. Supports 2 operations: set, get.

Usage examples:
- Set background color: ppt_background(operation='set', path='presentation.pptx', slideIndex=0, color='#FFFFFF')
- Set background image: ppt_background(operation='set', path='presentation.pptx', slideIndex=0, imagePath='bg.png')
- Get background: ppt_background(operation='get', path='presentation.pptx', slideIndex=0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set': Set slide background (required params: path, slideIndex)
- 'get': Get slide background (required params: path, slideIndex)",
                @enum = new[] { "set", "get" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
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
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for set operation, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        return operation.ToLower() switch
        {
            "set" => await SetBackgroundAsync(arguments, path),
            "get" => await GetBackgroundAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Sets slide background
    /// </summary>
    /// <param name="arguments">JSON arguments containing optional slideIndex, imagePath, color, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Success message</returns>
    private async Task<string> SetBackgroundAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", 0);
        var colorHex = ArgumentHelper.GetStringNullable(arguments, "color");
        var imagePath = ArgumentHelper.GetStringNullable(arguments, "imagePath");

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
            var color = ColorHelper.ParseColor(colorHex);
            fillFormat.FillType = FillType.Solid;
            fillFormat.SolidFillColor.Color = color;
        }
        else
        {
            throw new ArgumentException("Please provide at least one of color or imagePath");
        }

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"Background updated for slide {slideIndex}: {outputPath}");
    }

    /// <summary>
    /// Gets background information for a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing slideIndex</param>
    /// <param name="path">PowerPoint file path</param>
    /// <returns>Formatted string with background details</returns>
    private async Task<string> GetBackgroundAsync(JsonObject? arguments, string path)
    {
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

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

