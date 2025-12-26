using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint shape format (set, get)
///     Merges: PptSetShapeFormatTool, PptGetShapeFormatTool
/// </summary>
public class PptShapeFormatTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint shape format. Supports 2 operations: set, get.

Usage examples:
- Set format: ppt_shape_format(operation='set', path='presentation.pptx', slideIndex=0, shapeIndex=0, fillColor='#FF0000', lineColor='#0000FF')
- Get format: ppt_shape_format(operation='get', path='presentation.pptx', slideIndex=0, shapeIndex=0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set': Set shape format (required params: path, slideIndex, shapeIndex)
- 'get': Get shape format (required params: path, slideIndex, shapeIndex)",
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
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for set operation, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex", "shapeIndex" }
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
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

        return operation.ToLower() switch
        {
            "set" => await SetShapeFormatAsync(path, outputPath, slideIndex, shapeIndex, arguments),
            "get" => await GetShapeFormatAsync(path, slideIndex, shapeIndex),
            "set_line" => await SetShapeLineAsync(path, outputPath, slideIndex, shapeIndex, arguments),
            "set_fill" => await SetShapeFillAsync(path, outputPath, slideIndex, shapeIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets shape format properties
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="shapeIndex">Shape index (0-based)</param>
    /// <param name="arguments">JSON arguments containing optional fillColor, lineColor, lineWidth</param>
    /// <returns>Success message</returns>
    private Task<string> SetShapeFormatAsync(string path, string outputPath, int slideIndex, int shapeIndex,
        JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var x = ArgumentHelper.GetFloatNullable(arguments, "x");
            var y = ArgumentHelper.GetFloatNullable(arguments, "y");
            var width = ArgumentHelper.GetFloatNullable(arguments, "width");
            var height = ArgumentHelper.GetFloatNullable(arguments, "height");
            var rotation = ArgumentHelper.GetFloatNullable(arguments, "rotation");
            var fillColor = ArgumentHelper.GetStringNullable(arguments, "fillColor");
            var lineColor = ArgumentHelper.GetStringNullable(arguments, "lineColor");

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
                var color = ColorHelper.ParseColor(fillColor);
                shape.FillFormat.FillType = FillType.Solid;
                shape.FillFormat.SolidFillColor.Color = color;
            }

            if (!string.IsNullOrWhiteSpace(lineColor))
            {
                var color = ColorHelper.ParseColor(lineColor);
                shape.LineFormat.FillFormat.FillType = FillType.Solid;
                shape.LineFormat.FillFormat.SolidFillColor.Color = color;
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape format updated: slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets shape format information
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="shapeIndex">Shape index (0-based)</param>
    /// <returns>JSON string with shape format details</returns>
    private Task<string> GetShapeFormatAsync(string path, int slideIndex, int shapeIndex)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            // Fill format
            string? fillColorHex = null;
            if (shape.FillFormat.FillType == FillType.Solid)
            {
                var color = shape.FillFormat.SolidFillColor.Color;
                fillColorHex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }

            // Line format
            string? lineColorHex = null;
            if (shape.LineFormat.FillFormat.FillType == FillType.Solid)
            {
                var color = shape.LineFormat.FillFormat.SolidFillColor.Color;
                lineColorHex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }

            // Text format (if applicable)
            object? textFormat = null;
            if (shape is IAutoShape { TextFrame: not null } autoShape)
            {
                string? fontName = null;
                float? fontSize = null;
                bool? isBold = null;
                bool? isItalic = null;

                if (autoShape.TextFrame.Paragraphs.Count > 0)
                {
                    var firstPara = autoShape.TextFrame.Paragraphs[0];
                    if (firstPara.Portions.Count > 0)
                    {
                        var portion = firstPara.Portions[0];
                        fontName = portion.PortionFormat.LatinFont?.FontName;
                        fontSize = portion.PortionFormat.FontHeight;
                        isBold = portion.PortionFormat.FontBold == NullableBool.True;
                        isItalic = portion.PortionFormat.FontItalic == NullableBool.True;
                    }
                }

                textFormat = new
                {
                    text = autoShape.TextFrame.Text,
                    fontName,
                    fontSize,
                    isBold,
                    isItalic
                };
            }

            // Hyperlink
            string? hyperlink = null;
            if (shape.HyperlinkClick != null)
                hyperlink = shape.HyperlinkClick.ExternalUrl ?? (shape.HyperlinkClick.TargetSlide != null
                    ? $"Slide {presentation.Slides.IndexOf(shape.HyperlinkClick.TargetSlide)}"
                    : "Internal link");

            var result = new
            {
                slideIndex,
                shapeIndex,
                type = shape.GetType().Name,
                position = new { x = shape.X, y = shape.Y },
                size = new { width = shape.Width, height = shape.Height },
                rotation = shape.Rotation,
                fill = new
                {
                    type = shape.FillFormat.FillType.ToString(),
                    color = fillColorHex
                },
                line = new
                {
                    width = shape.LineFormat.Width,
                    color = lineColorHex
                },
                textFormat,
                hyperlink
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Sets shape line properties
    /// </summary>
    private Task<string> SetShapeLineAsync(string path, string outputPath, int slideIndex, int shapeIndex,
        JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var lineColor = ArgumentHelper.GetStringNullable(arguments, "color");
            var lineWidth = ArgumentHelper.GetFloatNullable(arguments, "width");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            if (!string.IsNullOrWhiteSpace(lineColor))
            {
                var color = ColorHelper.ParseColor(lineColor);
                shape.LineFormat.FillFormat.FillType = FillType.Solid;
                shape.LineFormat.FillFormat.SolidFillColor.Color = color;
            }

            if (lineWidth.HasValue) shape.LineFormat.Width = lineWidth.Value;

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape line format updated: slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets shape fill properties
    /// </summary>
    private Task<string> SetShapeFillAsync(string path, string outputPath, int slideIndex, int shapeIndex,
        JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fillType = ArgumentHelper.GetStringNullable(arguments, "fillType") ?? "Solid";
            var color = ArgumentHelper.GetStringNullable(arguments, "color");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            if (fillType.Equals("Solid", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(color))
            {
                var fillColor = ColorHelper.ParseColor(color);
                shape.FillFormat.FillType = FillType.Solid;
                shape.FillFormat.SolidFillColor.Color = fillColor;
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape fill format updated: slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }
}