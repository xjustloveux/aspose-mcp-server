using System.Text;
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
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "edit" => await EditShapeAsync(arguments, path, slideIndex),
            "delete" => await DeleteShapeAsync(arguments, path, slideIndex),
            "get" => await GetShapesAsync(arguments, path, slideIndex),
            "get_details" => await GetShapeDetailsAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Edits shape properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing shapeIndex, optional x, y, width, height, text, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> EditShapeAsync(JsonObject? arguments, string path, int slideIndex)
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
            var shapeDescription = targetShape.Placeholder != null
                ? $"Placeholder:{targetShape.Placeholder.Type}"
                : "NormalShape";
            var changes = new List<string>();
            if (x.HasValue)
            {
                targetShape.X = x.Value;
                changes.Add($"X: {x.Value}");
            }

            if (y.HasValue)
            {
                targetShape.Y = y.Value;
                changes.Add($"Y: {y.Value}");
            }

            if (width.HasValue)
            {
                targetShape.Width = width.Value;
                changes.Add($"Width: {width.Value}");
            }

            if (height.HasValue)
            {
                targetShape.Height = height.Value;
                changes.Add($"Height: {height.Value}");
            }

            if (rotation.HasValue)
            {
                targetShape.Rotation = rotation.Value;
                changes.Add($"Rotation: {rotation.Value}°");
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
                    changes.Add($"FlipHorizontal: {flipHorizontal.Value}");
                if (flipVertical.HasValue)
                    changes.Add($"FlipVertical: {flipVertical.Value}");
            }

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            presentation.Save(outputPath, SaveFormat.Pptx);

            return
                $"Shape {shapeIndex} ({shapeDescription}) edited on slide {slideIndex}: {string.Join(", ", changes)} - {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a shape from a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing shapeIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteShapeAsync(JsonObject? arguments, string path, int slideIndex)
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

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            presentation.Save(outputPath, SaveFormat.Pptx);
            return
                $"Shape {shapeIndex} deleted. Remaining non-placeholder shapes: {slide.Shapes.Count(s => s.Placeholder == null)}";
        });
    }

    /// <summary>
    ///     Gets all shapes from a slide
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Formatted string with all shapes</returns>
    private Task<string> GetShapesAsync(JsonObject? _, string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var sb = new StringBuilder();
            sb.AppendLine($"Slide {slideIndex} shapes: {slide.Shapes.Count}");

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
                var isPlaceholder = s.Placeholder != null ? " [PLACEHOLDER]" : "";
                sb.AppendLine(
                    $"[{i}] {kind}{isPlaceholder} pos=({s.X},{s.Y}) size=({s.Width},{s.Height}) text={(string.IsNullOrWhiteSpace(text) ? "(none)" : text)}");
            }

            return sb.ToString();
        });
    }

    /// <summary>
    ///     Gets detailed information about a specific shape
    /// </summary>
    /// <param name="arguments">JSON arguments containing shapeIndex</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Formatted string with shape details</returns>
    private Task<string> GetShapeDetailsAsync(JsonObject? arguments, string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            var sb = new StringBuilder();

            sb.AppendLine($"=== Shape {shapeIndex} Details ===");
            sb.AppendLine($"Type: {shape.GetType().Name}");
            sb.AppendLine($"IsPlaceholder: {shape.Placeholder != null}");
            sb.AppendLine($"Position: ({shape.X}, {shape.Y})");
            sb.AppendLine($"Size: ({shape.Width}, {shape.Height})");
            sb.AppendLine($"Rotation: {shape.Rotation}°");

            // Flip properties are accessed through Frame
            var frame = shape.Frame;
            sb.AppendLine($"FlipH: {frame.FlipH}");
            sb.AppendLine($"FlipV: {frame.FlipV}");

            if (shape is IAutoShape autoShape)
            {
                sb.AppendLine("\nAutoShape Properties:");
                sb.AppendLine($"  ShapeType: {autoShape.ShapeType}");
                sb.AppendLine($"  Text: {autoShape.TextFrame?.Text ?? "(none)"}");
                if (autoShape.HyperlinkClick != null)
                {
                    var url = autoShape.HyperlinkClick.ExternalUrl ?? (autoShape.HyperlinkClick.TargetSlide != null
                        ? $"Slide {presentation.Slides.IndexOf(autoShape.HyperlinkClick.TargetSlide)}"
                        : "Internal link");
                    sb.AppendLine($"  Hyperlink: {url}");
                }
            }
            else if (shape is PictureFrame picture)
            {
                sb.AppendLine("\nPicture Properties:");
                sb.AppendLine($"  AlternativeText: {picture.AlternativeText ?? "(none)"}");
            }
            else if (shape is ITable table)
            {
                sb.AppendLine("\nTable Properties:");
                sb.AppendLine($"  Rows: {table.Rows.Count}");
                sb.AppendLine($"  Columns: {table.Columns.Count}");
            }
            else if (shape is IChart chart)
            {
                sb.AppendLine("\nChart Properties:");
                sb.AppendLine($"  ChartType: {chart.Type}");
                sb.AppendLine($"  Title: {chart.ChartTitle?.TextFrameForOverriding?.Text ?? "(none)"}");
            }

            return sb.ToString();
        });
    }
}