using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint backgrounds (set, get)
///     Merges: PptSetBackgroundTool, PptGetBackgroundTool
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

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "set" => await SetBackgroundAsync(path, outputPath, arguments),
            "get" => await GetBackgroundAsync(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets slide background
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing optional slideIndex, imagePath, color</param>
    /// <returns>Success message</returns>
    private Task<string> SetBackgroundAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(async () =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", 0);
            var colorHex = ArgumentHelper.GetStringNullable(arguments, "color");
            var imagePath = ArgumentHelper.GetStringNullable(arguments, "imagePath");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var fillFormat = slide.Background.FillFormat;

            if (!string.IsNullOrWhiteSpace(imagePath))
            {
                var img = presentation.Images.AddImage(await File.ReadAllBytesAsync(imagePath));
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

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Background updated for slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets background information for a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex</param>
    /// <returns>JSON string with background details</returns>
    private Task<string> GetBackgroundAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var background = slide.Background;

            string? colorHex = null;
            if (background?.FillFormat.FillType == FillType.Solid)
            {
                var color = background.FillFormat.SolidFillColor.Color;
                colorHex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }

            var result = new
            {
                slideIndex,
                hasBackground = background != null,
                fillType = background?.FillFormat.FillType.ToString(),
                color = colorHex,
                isPictureFill = background?.FillFormat.FillType == FillType.Picture
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}