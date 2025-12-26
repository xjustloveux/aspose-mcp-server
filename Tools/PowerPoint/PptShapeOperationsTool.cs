using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint shape operations (group, ungroup, copy, reorder, align, flip)
///     Merges: PptGroupShapesTool, PptUngroupShapesTool, PptCopyShapeTool, PptReorderShapeTool,
///     PptAlignShapesTool, PptFlipShapeTool
/// </summary>
public class PptShapeOperationsTool : IAsposeTool
{
    public string Description =>
        @"PowerPoint shape operations. Supports 6 operations: group, ungroup, copy, reorder, align, flip.

Usage examples:
- Group shapes: ppt_shape_operations(operation='group', path='presentation.pptx', slideIndex=0, shapeIndices=[0,1,2])
- Ungroup shape: ppt_shape_operations(operation='ungroup', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Copy shape: ppt_shape_operations(operation='copy', path='presentation.pptx', slideIndex=0, shapeIndex=0, fromSlide=0, toSlide=1)
- Reorder shape: ppt_shape_operations(operation='reorder', path='presentation.pptx', slideIndex=0, shapeIndex=0, newIndex=2)
- Align shapes: ppt_shape_operations(operation='align', path='presentation.pptx', slideIndex=0, shapeIndices=[0,1,2], alignment='left')
- Flip shape: ppt_shape_operations(operation='flip', path='presentation.pptx', slideIndex=0, shapeIndex=0, flipType='horizontal')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'group': Group multiple shapes (required params: path, slideIndex, shapeIndices)
- 'ungroup': Ungroup a shape (required params: path, slideIndex, shapeIndex)
- 'copy': Copy shape to another slide (required params: path, slideIndex, shapeIndex, fromSlide, toSlide)
- 'reorder': Reorder shape position (required params: path, slideIndex, shapeIndex, newIndex)
- 'align': Align multiple shapes (required params: path, slideIndex, shapeIndices, alignment)
- 'flip': Flip a shape (required params: path, slideIndex, shapeIndex, flipType)",
                @enum = new[] { "group", "ungroup", "copy", "reorder", "align", "flip" }
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
                description = "Shape index (0-based, required for ungroup/copy/reorder/flip)"
            },
            shapeIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Shape indices array (required for group/align)"
            },
            fromSlide = new
            {
                type = "number",
                description = "Source slide index (0-based, required for copy)"
            },
            toSlide = new
            {
                type = "number",
                description = "Target slide index (0-based, required for copy)"
            },
            toIndex = new
            {
                type = "number",
                description = "Target index (0-based, required for reorder)"
            },
            align = new
            {
                type = "string",
                description = "Alignment: left|center|right|top|middle|bottom (required for align)"
            },
            alignToSlide = new
            {
                type = "boolean",
                description = "Align to slide instead of group (optional, for align, default: false)"
            },
            flipHorizontal = new
            {
                type = "boolean",
                description = "Flip horizontally (optional, for flip)"
            },
            flipVertical = new
            {
                type = "boolean",
                description = "Flip vertically (optional, for flip)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for all operations, defaults to input path)"
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
            "group" => await GroupShapesAsync(path, outputPath, slideIndex, arguments),
            "ungroup" => await UngroupShapesAsync(path, outputPath, slideIndex, arguments),
            "copy" => await CopyShapeAsync(path, outputPath, arguments),
            "reorder" => await ReorderShapeAsync(path, outputPath, slideIndex, arguments),
            "align" => await AlignShapesAsync(path, outputPath, slideIndex, arguments),
            "flip" => await FlipShapeAsync(path, outputPath, slideIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Groups shapes together
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing shapeIndexes array</param>
    /// <returns>Success message</returns>
    private Task<string> GroupShapesAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            try
            {
                var shapeIndicesArray = ArgumentHelper.GetArray(arguments, "shapeIndices");

                var shapeIndices = shapeIndicesArray.Select(s => s?.GetValue<int>()).Where(s => s.HasValue)
                    .Select(s => s!.Value).OrderByDescending(s => s).ToList();

                if (shapeIndices.Count < 2) throw new ArgumentException("At least 2 shapes are required for grouping");

                using var presentation = new Presentation(path);
                if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
                    throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

                var slide = presentation.Slides[slideIndex];
                var shapesToGroup = new List<(IShape shape, int originalIndex)>();

                // Use unified index logic: directly access slide.Shapes[index]
                foreach (var idx in shapeIndices.OrderByDescending(x => x))
                {
                    if (idx < 0 || idx >= slide.Shapes.Count)
                        throw new ArgumentException($"shapeIndex {idx} is out of range (0-{slide.Shapes.Count - 1})");
                    var shape = slide.Shapes[idx];
                    shapesToGroup.Add((shape, idx));
                }

                // Calculate bounding box for all shapes
                var minX = shapesToGroup.Min(s => s.shape.X);
                var minY = shapesToGroup.Min(s => s.shape.Y);
                var maxX = shapesToGroup.Max(s => s.shape.X + s.shape.Width);
                var maxY = shapesToGroup.Max(s => s.shape.Y + s.shape.Height);
                var groupWidth = maxX - minX;
                var groupHeight = maxY - minY;

                try
                {
                    // Group shapes - create a group shape with calculated bounds and add shapes to it
                    var groupShape = slide.Shapes.AddGroupShape();
                    groupShape.X = minX;
                    groupShape.Y = minY;
                    groupShape.Width = groupWidth;
                    groupShape.Height = groupHeight;

                    // Remove shapes from slide (in reverse order to maintain indices) and add to group
                    foreach (var (shape, originalIndex) in shapesToGroup)
                    {
                        // Calculate relative position within the group
                        var relativeX = shape.X - minX;
                        var relativeY = shape.Y - minY;

                        // Clone shape to group
                        var clonedShape = groupShape.Shapes.AddClone(shape);
                        clonedShape.X = relativeX;
                        clonedShape.Y = relativeY;

                        // Remove original shape from slide (using originalIndex which is already in descending order)
                        slide.Shapes.RemoveAt(originalIndex);
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to group shapes: {ex.Message}", ex);
                }

                presentation.Save(outputPath, SaveFormat.Pptx);
                return $"Grouped {shapeIndices.Count} shapes on slide {slideIndex}. Output: {outputPath}";
            }
            catch (ArgumentException)
            {
                throw; // Re-throw ArgumentException as-is
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error grouping shapes: {ex.Message}", ex);
            }
        });
    }

    /// <summary>
    ///     Ungroups shapes
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing shapeIndex</param>
    /// <returns>Success message</returns>
    private Task<string> UngroupShapesAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            if (shape is not IGroupShape groupShape)
                throw new ArgumentException($"Shape at index {shapeIndex} is not a group");

            // Ungroup - add shapes back to slide and remove group
            var shapesInGroup = new List<IShape>();
            foreach (var s in groupShape.Shapes) shapesInGroup.Add(s);

            foreach (var s in shapesInGroup) slide.Shapes.AddClone(s);

            slide.Shapes.Remove(groupShape);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Ungrouped shape on slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Copies a shape to another slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sourceSlideIndex, sourceShapeIndex, targetSlideIndex</param>
    /// <returns>Success message</returns>
    private Task<string> CopyShapeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fromSlide = ArgumentHelper.GetInt(arguments, "fromSlide");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var toSlide = ArgumentHelper.GetInt(arguments, "toSlide");

            using var presentation = new Presentation(path);
            if (fromSlide < 0 || fromSlide >= presentation.Slides.Count)
                throw new ArgumentException("fromSlide out of range");
            if (toSlide < 0 || toSlide >= presentation.Slides.Count)
                throw new ArgumentException("toSlide out of range");

            var sourceSlide = presentation.Slides[fromSlide];
            if (shapeIndex < 0 || shapeIndex >= sourceSlide.Shapes.Count)
                throw new ArgumentException("shapeIndex out of range");

            var targetSlide = presentation.Slides[toSlide];
            targetSlide.Shapes.AddClone(sourceSlide.Shapes[shapeIndex]);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape {shapeIndex} copied from slide {fromSlide} to slide {toSlide}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Changes the z-order of a shape
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, newIndex</param>
    /// <returns>Success message</returns>
    private Task<string> ReorderShapeAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var toIndex = ArgumentHelper.GetInt(arguments, "toIndex");

            using var presentation = new Presentation(path);
            if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

            var slide = presentation.Slides[slideIndex];
            var count = slide.Shapes.Count;
            if (shapeIndex < 0 || shapeIndex >= count)
                throw new ArgumentException($"shapeIndex must be between 0 and {count - 1}");
            if (toIndex < 0 || toIndex >= count)
                throw new ArgumentException($"toIndex must be between 0 and {count - 1}");

            var shape = slide.Shapes[shapeIndex];
            slide.Shapes.InsertClone(toIndex, shape);
            var removeIndex = shapeIndex + (shapeIndex < toIndex ? 1 : 0);
            slide.Shapes.RemoveAt(removeIndex);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape moved: {shapeIndex} -> {toIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Aligns shapes
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing shapeIndexes array, alignmentType</param>
    /// <returns>Success message</returns>
    private Task<string> AlignShapesAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var alignStr = ArgumentHelper.GetString(arguments, "align");
            var shapeIndicesArray = ArgumentHelper.GetArray(arguments, "shapeIndices");

            var shapeIndices = shapeIndicesArray.Select(x => x?.GetValue<int>() ?? -1).ToArray();
            var alignToSlide = ArgumentHelper.GetBool(arguments, "alignToSlide", false);

            if (shapeIndices.Length < 2) throw new ArgumentException("shapeIndices must contain at least 2 items");

            using var presentation = new Presentation(path);
            if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

            var slide = presentation.Slides[slideIndex];
            foreach (var idx in shapeIndices)
                if (idx < 0 || idx >= slide.Shapes.Count)
                    throw new ArgumentException($"shape index {idx} is out of range (0-{slide.Shapes.Count - 1})");

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

            presentation.Save(outputPath, SaveFormat.Pptx);
            return
                $"Aligned {shapeIndices.Length} shapes: {alignStr}, alignToSlide={alignToSlide}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Flips a shape horizontally or vertically
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, flipType (horizontal/vertical)</param>
    /// <returns>Success message</returns>
    private Task<string> FlipShapeAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var flipHorizontal = ArgumentHelper.GetBoolNullable(arguments, "flipHorizontal");
            var flipVertical = ArgumentHelper.GetBoolNullable(arguments, "flipVertical");

            if (!flipHorizontal.HasValue && !flipVertical.HasValue)
                throw new ArgumentException("At least one of flipHorizontal or flipVertical must be provided");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            // Apply flip transformations using Frame property
            if (flipHorizontal.HasValue || flipVertical.HasValue)
            {
                var currentFrame = shape.Frame;
                var newFlipH = flipHorizontal.HasValue
                    ? flipHorizontal.Value ? NullableBool.True : NullableBool.False
                    : currentFrame.FlipH;
                var newFlipV = flipVertical.HasValue
                    ? flipVertical.Value ? NullableBool.True : NullableBool.False
                    : currentFrame.FlipV;

                // Create a new ShapeFrame with updated flip properties
                shape.Frame = new ShapeFrame(
                    currentFrame.X,
                    currentFrame.Y,
                    currentFrame.Width,
                    currentFrame.Height,
                    newFlipH,
                    newFlipV,
                    currentFrame.Rotation
                );
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape flipped on slide {slideIndex}, shape {shapeIndex}. Output: {outputPath}";
        });
    }
}