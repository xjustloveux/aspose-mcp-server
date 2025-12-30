using System.Drawing;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint backgrounds (set, get).
/// </summary>
public class PptBackgroundTool : IAsposeTool
{
    /// <inheritdoc />
    public string Description => @"Manage PowerPoint backgrounds. Supports 2 operations: set, get.

Usage examples:
- Set background color: ppt_background(operation='set', path='presentation.pptx', slideIndex=0, color='#FFFFFF')
- Set background image: ppt_background(operation='set', path='presentation.pptx', slideIndex=0, imagePath='bg.png')
- Apply to all slides: ppt_background(operation='set', path='presentation.pptx', color='#FFFFFF', applyToAll=true)
- Get background: ppt_background(operation='get', path='presentation.pptx', slideIndex=0)";

    /// <inheritdoc />
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set': Set slide background
- 'get': Get slide background info",
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
                description = "Slide index (0-based, default: 0, ignored if applyToAll is true)"
            },
            color = new
            {
                type = "string",
                description = "Hex color like #FFAA00 or #80FFAA00 (with alpha)"
            },
            imagePath = new
            {
                type = "string",
                description = "Background image path"
            },
            applyToAll = new
            {
                type = "boolean",
                description = "Apply background to all slides (default: false)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <inheritdoc />
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
    ///     Sets slide background with color or image.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex, color, imagePath, applyToAll.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when neither color nor imagePath is provided.</exception>
    private Task<string> SetBackgroundAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(async () =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", 0);
            var colorHex = ArgumentHelper.GetStringNullable(arguments, "color");
            var imagePath = ArgumentHelper.GetStringNullable(arguments, "imagePath");
            var applyToAll = ArgumentHelper.GetBool(arguments, "applyToAll", false);

            if (string.IsNullOrWhiteSpace(colorHex) && string.IsNullOrWhiteSpace(imagePath))
                throw new ArgumentException("Please provide at least one of color or imagePath");

            using var presentation = new Presentation(path);

            IPPImage? img = null;
            if (!string.IsNullOrWhiteSpace(imagePath))
                img = presentation.Images.AddImage(await File.ReadAllBytesAsync(imagePath));

            Color? color = null;
            if (!string.IsNullOrWhiteSpace(colorHex))
                color = ColorHelper.ParseColor(colorHex);

            var slidesToUpdate = applyToAll
                ? presentation.Slides.ToList()
                : [PowerPointHelper.GetSlide(presentation, slideIndex)];

            foreach (var slide in slidesToUpdate)
                ApplyBackground(slide, color, img);

            presentation.Save(outputPath, SaveFormat.Pptx);

            var message = applyToAll
                ? $"Background applied to all {slidesToUpdate.Count} slides"
                : $"Background updated for slide {slideIndex}";
            return $"{message}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Applies background color or image to a slide.
    /// </summary>
    /// <param name="slide">Target slide.</param>
    /// <param name="color">Background color (nullable).</param>
    /// <param name="image">Background image (nullable).</param>
    private static void ApplyBackground(ISlide slide, Color? color, IPPImage? image)
    {
        slide.Background.Type = BackgroundType.OwnBackground;
        var fillFormat = slide.Background.FillFormat;

        if (image != null)
        {
            fillFormat.FillType = FillType.Picture;
            fillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
            fillFormat.PictureFillFormat.Picture.Image = image;
        }
        else if (color.HasValue)
        {
            fillFormat.FillType = FillType.Solid;
            fillFormat.SolidFillColor.Color = color.Value;
        }
    }

    /// <summary>
    ///     Gets background information for a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="arguments">JSON arguments containing slideIndex.</param>
    /// <returns>JSON string with background details including color and opacity.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range.</exception>
    private Task<string> GetBackgroundAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex", 0);

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var background = slide.Background;
            var fillFormat = background?.FillFormat;

            string? colorHex = null;
            double? opacity = null;

            if (fillFormat?.FillType == FillType.Solid)
                try
                {
                    var solidColor = fillFormat.SolidFillColor.Color;
                    if (!solidColor.IsEmpty)
                    {
                        colorHex = solidColor.A < 255
                            ? $"#{solidColor.A:X2}{solidColor.R:X2}{solidColor.G:X2}{solidColor.B:X2}"
                            : $"#{solidColor.R:X2}{solidColor.G:X2}{solidColor.B:X2}";
                        opacity = Math.Round(solidColor.A / 255.0, 2);
                    }
                }
                catch
                {
                    // Theme colors may throw exceptions, return null for color
                }

            var result = new
            {
                slideIndex,
                hasBackground = background != null,
                fillType = fillFormat?.FillType.ToString(),
                color = colorHex,
                opacity,
                isPictureFill = fillFormat?.FillType == FillType.Picture
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}