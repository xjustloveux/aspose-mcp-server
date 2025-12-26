using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint shapes (edit, delete, get, get details)
///     Merges: PptEditShapeTool, PptDeleteShapeTool, PptGetShapesTool, PptGetShapeDetailsTool
/// </summary>
public class PptShapeTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint shapes. Supports 4 operations: edit, delete, get, get_details.

Usage examples:
- Edit shape: ppt_shape(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, x=200, y=200)
- Delete shape: ppt_shape(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get shapes: ppt_shape(operation='get', path='presentation.pptx', slideIndex=0)
- Get details: ppt_shape(operation='get_details', path='presentation.pptx', slideIndex=0, shapeIndex=0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'edit': Edit shape properties (required params: path, slideIndex, shapeIndex)
- 'delete': Delete a shape (required params: path, slideIndex, shapeIndex)
- 'get': Get all shapes (required params: path, slideIndex)
- 'get_details': Get shape details (required params: path, slideIndex, shapeIndex)",
                @enum = new[] { "edit", "delete", "get", "get_details" }
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
                description =
                    "Shape index (0-based, required for edit/delete/get_details). Refers to the index among non-placeholder shapes on the slide."
            },
            x = new
            {
                type = "number",
                description = "X position (optional, for edit)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, for edit)"
            },
            width = new
            {
                type = "number",
                description = "Width (optional, for edit)"
            },
            height = new
            {
                type = "number",
                description = "Height (optional, for edit)"
            },
            rotation = new
            {
                type = "number",
                description = "Rotation angle in degrees (optional, for edit)"
            },
            flipHorizontal = new
            {
                type = "boolean",
                description = "Flip horizontally (optional, for edit)"
            },
            flipVertical = new
            {
                type = "boolean",
                description = "Flip vertically (optional, for edit)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
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

        return operation.ToLower() switch
        {
            "edit" => await EditShapeAsync(path, outputPath, slideIndex, arguments),
            "delete" => await DeleteShapeAsync(path, outputPath, slideIndex, arguments),
            "get" => await GetShapesAsync(path, slideIndex),
            "get_details" => await GetShapeDetailsAsync(path, slideIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Edits shape properties
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, optional x, y, width, height</param>
    /// <returns>Success message</returns>
    private Task<string> EditShapeAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var x = ArgumentHelper.GetFloatNullable(arguments, "x");
            var y = ArgumentHelper.GetFloatNullable(arguments, "y");
            var width = ArgumentHelper.GetFloatNullable(arguments, "width");
            var height = ArgumentHelper.GetFloatNullable(arguments, "height");
            var rotation = ArgumentHelper.GetFloatNullable(arguments, "rotation");
            var flipHorizontal = ArgumentHelper.GetBoolNullable(arguments, "flipHorizontal");
            var flipVertical = ArgumentHelper.GetBoolNullable(arguments, "flipVertical");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var nonPlaceholderShapes = slide.Shapes.Where(s => s.Placeholder == null).ToList();

            if (shapeIndex < 0 || shapeIndex >= nonPlaceholderShapes.Count)
            {
                var totalShapes = slide.Shapes.Count;
                var totalNonPlaceholder = nonPlaceholderShapes.Count;
                throw new ArgumentException(
                    $"Index {shapeIndex} is out of range for user-defined shapes. " +
                    $"Slide {slideIndex} has {totalNonPlaceholder} non-placeholder shape(s) " +
                    $"(out of {totalShapes} total shape(s)). Valid indices: 0 to {totalNonPlaceholder - 1}.");
            }

            var targetShape = nonPlaceholderShapes[shapeIndex];
            var changes = new List<string>();
            if (x.HasValue)
            {
                targetShape.X = x.Value;
                changes.Add($"X={x.Value}");
            }

            if (y.HasValue)
            {
                targetShape.Y = y.Value;
                changes.Add($"Y={y.Value}");
            }

            if (width.HasValue)
            {
                targetShape.Width = width.Value;
                changes.Add($"Width={width.Value}");
            }

            if (height.HasValue)
            {
                targetShape.Height = height.Value;
                changes.Add($"Height={height.Value}");
            }

            if (rotation.HasValue)
            {
                targetShape.Rotation = rotation.Value;
                changes.Add($"Rotation={rotation.Value}");
            }

            if (flipHorizontal.HasValue || flipVertical.HasValue)
            {
                var currentFrame = targetShape.Frame;
                var newFlipH = flipHorizontal.HasValue
                    ? flipHorizontal.Value ? NullableBool.True : NullableBool.False
                    : currentFrame.FlipH;
                var newFlipV = flipVertical.HasValue
                    ? flipVertical.Value ? NullableBool.True : NullableBool.False
                    : currentFrame.FlipV;

                targetShape.Frame = new ShapeFrame(
                    targetShape.X,
                    targetShape.Y,
                    targetShape.Width,
                    targetShape.Height,
                    newFlipH,
                    newFlipV,
                    targetShape.Rotation
                );

                if (flipHorizontal.HasValue)
                    changes.Add($"FlipH={flipHorizontal.Value}");
                if (flipVertical.HasValue)
                    changes.Add($"FlipV={flipVertical.Value}");
            }

            presentation.Save(outputPath, SaveFormat.Pptx);

            return
                $"Shape {shapeIndex} on slide {slideIndex} edited ({string.Join(", ", changes)}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a shape from a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing shapeIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteShapeAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var nonPlaceholderShapes = slide.Shapes.Where(s => s.Placeholder == null).ToList();

            if (shapeIndex < 0 || shapeIndex >= nonPlaceholderShapes.Count)
            {
                var totalShapes = slide.Shapes.Count;
                var totalNonPlaceholder = nonPlaceholderShapes.Count;
                throw new ArgumentException(
                    $"Cannot delete: Shape index {shapeIndex} is out of range. " +
                    $"Slide {slideIndex} has {totalNonPlaceholder} non-placeholder shape(s) " +
                    $"(out of {totalShapes} total shape(s)). Valid indices: 0 to {totalNonPlaceholder - 1}.");
            }

            var shapeToDelete = nonPlaceholderShapes[shapeIndex];

            var originalIndex = -1;
            for (var i = 0; i < slide.Shapes.Count; i++)
                if (slide.Shapes[i] == shapeToDelete)
                {
                    originalIndex = i;
                    break;
                }

            if (originalIndex >= 0 && originalIndex < slide.Shapes.Count)
                slide.Shapes.RemoveAt(originalIndex);
            else
                slide.Shapes.Remove(shapeToDelete);

            presentation.Save(outputPath, SaveFormat.Pptx);
            var remaining = slide.Shapes.Count(s => s.Placeholder == null);
            return $"Shape {shapeIndex} on slide {slideIndex} deleted ({remaining} remaining). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all shapes from a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>JSON string with all shapes</returns>
    private Task<string> GetShapesAsync(string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var shapesList = new List<object>();
            for (var i = 0; i < slide.Shapes.Count; i++)
            {
                var s = slide.Shapes[i];
                var kind = s switch
                {
                    IAutoShape => "AutoShape",
                    PictureFrame => "Picture",
                    ITable => "Table",
                    IChart => "Chart",
                    IGroupShape => "Group",
                    ISmartArt => "SmartArt",
                    IAudioFrame => "Audio",
                    IVideoFrame => "Video",
                    _ => s.GetType().Name
                };

                var text = (s as IAutoShape)?.TextFrame?.Text;

                shapesList.Add(new
                {
                    index = i,
                    type = kind,
                    isPlaceholder = s.Placeholder != null,
                    position = new { x = s.X, y = s.Y },
                    size = new { width = s.Width, height = s.Height },
                    text = string.IsNullOrWhiteSpace(text) ? null : text
                });
            }

            var result = new
            {
                slideIndex,
                count = slide.Shapes.Count,
                shapes = shapesList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Gets detailed information about a specific shape
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing shapeIndex</param>
    /// <returns>JSON string with shape details</returns>
    private Task<string> GetShapeDetailsAsync(string path, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            // Flip properties are accessed through Frame
            var frame = shape.Frame;

            object? typeSpecificProperties = null;

            if (shape is IAutoShape autoShape)
            {
                string? hyperlink = null;
                if (autoShape.HyperlinkClick != null)
                    hyperlink = autoShape.HyperlinkClick.ExternalUrl ?? (autoShape.HyperlinkClick.TargetSlide != null
                        ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkClick.TargetSlide)}"
                        : "Internal link");

                typeSpecificProperties = new
                {
                    shapeType = autoShape.ShapeType.ToString(),
                    text = autoShape.TextFrame?.Text,
                    hyperlink
                };
            }
            else if (shape is PictureFrame picture)
            {
                typeSpecificProperties = new
                {
                    alternativeText = picture.AlternativeText
                };
            }
            else if (shape is ITable table)
            {
                typeSpecificProperties = new
                {
                    rows = table.Rows.Count,
                    columns = table.Columns.Count
                };
            }
            else if (shape is IChart chart)
            {
                typeSpecificProperties = new
                {
                    chartType = chart.Type.ToString(),
                    title = chart.ChartTitle?.TextFrameForOverriding?.Text
                };
            }

            var result = new
            {
                slideIndex,
                shapeIndex,
                type = shape.GetType().Name,
                isPlaceholder = shape.Placeholder != null,
                position = new { x = shape.X, y = shape.Y },
                size = new { width = shape.Width, height = shape.Height },
                rotation = shape.Rotation,
                flipH = frame.FlipH.ToString(),
                flipV = frame.FlipV.ToString(),
                properties = typeSpecificProperties
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }
}