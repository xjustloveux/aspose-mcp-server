using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.ShapeDetailProviders;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint shapes.
///     Supports 12 operations: get, get_details, delete, edit, set_format, clear_format,
///     group, ungroup, copy, reorder, align, flip
/// </summary>
public class PptShapeTool : IAsposeTool
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        NumberHandling = JsonNumberHandling.AllowNamedFloatingPointLiterals
    };

    public string Description => @"Unified PowerPoint shape management tool. Supports 12 operations.

Note: Position/size units are in points (1 point = 1/72 inch).
Note: shapeIndex uses original slide.Shapes index. Use 'get' to see all shapes with their indices.

Operations:
- Basic: get, get_details, delete
- Edit: edit, set_format, clear_format
- Advanced: group, ungroup, copy, reorder, align, flip

Usage examples:
- Get shapes: ppt_shape(operation='get', path='file.pptx', slideIndex=0)
- Get details: ppt_shape(operation='get_details', path='file.pptx', slideIndex=0, shapeIndex=0)
- Delete: ppt_shape(operation='delete', path='file.pptx', slideIndex=0, shapeIndex=0)
- Edit: ppt_shape(operation='edit', path='file.pptx', slideIndex=0, shapeIndex=0, x=100, y=100)
- Set format: ppt_shape(operation='set_format', path='file.pptx', slideIndex=0, shapeIndex=0, fillColor='#FF0000')
- Clear format: ppt_shape(operation='clear_format', path='file.pptx', slideIndex=0, shapeIndex=0, clearFill=true)
- Group: ppt_shape(operation='group', path='file.pptx', slideIndex=0, shapeIndices=[0,1,2])
- Ungroup: ppt_shape(operation='ungroup', path='file.pptx', slideIndex=0, shapeIndex=0)
- Copy: ppt_shape(operation='copy', path='file.pptx', fromSlide=0, toSlide=1, shapeIndex=0)
- Reorder: ppt_shape(operation='reorder', path='file.pptx', slideIndex=0, shapeIndex=0, toIndex=2)
- Align: ppt_shape(operation='align', path='file.pptx', slideIndex=0, shapeIndices=[0,1], align='left')
- Flip: ppt_shape(operation='flip', path='file.pptx', slideIndex=0, shapeIndex=0, flipHorizontal=true)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform:
- 'get': List all shapes on slide
- 'get_details': Get detailed shape info
- 'delete': Delete a shape
- 'edit': Edit shape (position, size, rotation, text)
- 'set_format': Set fill/line format
- 'clear_format': Clear fill/line format
- 'group': Group multiple shapes
- 'ungroup': Ungroup a group shape
- 'copy': Copy shape to another slide
- 'reorder': Change shape Z-order
- 'align': Align multiple shapes
- 'flip': Flip shape horizontally/vertically",
                @enum = new[]
                {
                    "get", "get_details", "delete", "edit", "set_format", "clear_format", "group", "ungroup", "copy",
                    "reorder", "align", "flip"
                }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based, required except for copy)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index (0-based, for single-shape operations)"
            },
            shapeIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Shape indices array (for group/align)"
            },
            // Edit parameters
            x = new { type = "number", description = "X position in points (for edit)" },
            y = new { type = "number", description = "Y position in points (for edit)" },
            width = new { type = "number", description = "Width in points (for edit)" },
            height = new { type = "number", description = "Height in points (for edit)" },
            rotation = new { type = "number", description = "Rotation in degrees (for edit)" },
            text = new { type = "string", description = "Text content for AutoShape (for edit)" },
            // Format parameters
            fillColor = new { type = "string", description = "Fill color hex, e.g. #FF0000 (for set_format)" },
            lineColor = new { type = "string", description = "Line color hex (for set_format)" },
            lineWidth = new { type = "number", description = "Line width in points (for set_format)" },
            clearFill = new { type = "boolean", description = "Clear fill (for clear_format)" },
            clearLine = new { type = "boolean", description = "Clear line (for clear_format)" },
            // Copy parameters
            fromSlide = new { type = "number", description = "Source slide index (for copy)" },
            toSlide = new { type = "number", description = "Target slide index (for copy)" },
            // Reorder parameters
            toIndex = new { type = "number", description = "Target Z-order index (for reorder)" },
            // Align parameters
            align = new
            {
                type = "string",
                description = "Alignment: left|center|right|top|middle|bottom (for align)"
            },
            alignToSlide = new { type = "boolean", description = "Align to slide bounds (for align, default: false)" },
            // Flip parameters
            flipHorizontal = new { type = "boolean", description = "Flip horizontally (for flip/edit)" },
            flipVertical = new { type = "boolean", description = "Flip vertically (for flip/edit)" },
            // Output
            outputPath = new { type = "string", description = "Output file path (optional, defaults to input path)" }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            // Basic operations
            "get" => await GetShapesAsync(path, arguments),
            "get_details" => await GetShapeDetailsAsync(path, arguments),
            "delete" => await DeleteShapeAsync(path, outputPath, arguments),
            // Edit operations
            "edit" => await EditShapeAsync(path, outputPath, arguments),
            "set_format" => await SetFormatAsync(path, outputPath, arguments),
            "clear_format" => await ClearFormatAsync(path, outputPath, arguments),
            // Advanced operations
            "group" => await GroupShapesAsync(path, outputPath, arguments),
            "ungroup" => await UngroupShapesAsync(path, outputPath, arguments),
            "copy" => await CopyShapeAsync(path, outputPath, arguments),
            "reorder" => await ReorderShapeAsync(path, outputPath, arguments),
            "align" => await AlignShapesAsync(path, outputPath, arguments),
            "flip" => await FlipShapeAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    #region Basic Operations

    /// <summary>
    ///     Gets all shapes from a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex</param>
    /// <returns>JSON string with all shapes</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is out of range</exception>
    private Task<string> GetShapesAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var shapesList = new List<object>();
            var userShapeCounter = 0;
            for (var i = 0; i < slide.Shapes.Count; i++)
            {
                var s = slide.Shapes[i];
                var isPlaceholder = s.Placeholder != null;
                var (typeName, _) = ShapeDetailProviderFactory.GetShapeDetails(s, presentation);

                var text = (s as IAutoShape)?.TextFrame?.Text;

                shapesList.Add(new
                {
                    index = i,
                    userShapeIndex = isPlaceholder ? (int?)-1 : userShapeCounter,
                    type = typeName,
                    isPlaceholder,
                    position = new { x = s.X, y = s.Y },
                    size = new { width = s.Width, height = s.Height },
                    text = string.IsNullOrWhiteSpace(text) ? null : text
                });

                if (!isPlaceholder) userShapeCounter++;
            }

            var result = new
            {
                slideIndex,
                totalCount = slide.Shapes.Count,
                userShapeCount = userShapeCounter,
                shapes = shapesList
            };

            return JsonSerializer.Serialize(result, JsonOptions);
        });
    }

    /// <summary>
    ///     Gets detailed information about a specific shape
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex</param>
    /// <returns>JSON string with shape details including format info</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is out of range</exception>
    private Task<string> GetShapeDetailsAsync(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            var frame = shape.Frame;
            var (typeName, typeSpecificProperties) = ShapeDetailProviderFactory.GetShapeDetails(shape, presentation);

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

            var result = new
            {
                slideIndex,
                shapeIndex,
                type = typeName,
                isPlaceholder = shape.Placeholder != null,
                position = new { x = shape.X, y = shape.Y },
                size = new { width = shape.Width, height = shape.Height },
                rotation = shape.Rotation,
                flipH = frame.FlipH.ToString(),
                flipV = frame.FlipV.ToString(),
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
                properties = typeSpecificProperties
            };

            return JsonSerializer.Serialize(result, JsonOptions);
        });
    }

    /// <summary>
    ///     Deletes a shape from a slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is out of range</exception>
    private Task<string> DeleteShapeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            PowerPointHelper.ValidateCollectionIndex(shapeIndex, slide.Shapes.Count, "shapeIndex");

            slide.Shapes.RemoveAt(shapeIndex);

            presentation.Save(outputPath, SaveFormat.Pptx);
            var remaining = slide.Shapes.Count;
            return $"Shape {shapeIndex} deleted from slide {slideIndex} ({remaining} remaining). Output: {outputPath}";
        });
    }

    #endregion

    #region Edit Operations

    /// <summary>
    ///     Edits shape properties
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, and optional edit parameters</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is out of range</exception>
    private Task<string> EditShapeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var x = ArgumentHelper.GetFloatNullable(arguments, "x");
            var y = ArgumentHelper.GetFloatNullable(arguments, "y");
            var width = ArgumentHelper.GetFloatNullable(arguments, "width");
            var height = ArgumentHelper.GetFloatNullable(arguments, "height");
            var rotation = ArgumentHelper.GetFloatNullable(arguments, "rotation");
            var flipHorizontal = ArgumentHelper.GetBoolNullable(arguments, "flipHorizontal");
            var flipVertical = ArgumentHelper.GetBoolNullable(arguments, "flipVertical");
            var text = ArgumentHelper.GetStringNullable(arguments, "text");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            var changes = new List<string>();

            if (x.HasValue)
            {
                shape.X = x.Value;
                changes.Add($"X={x.Value}");
            }

            if (y.HasValue)
            {
                shape.Y = y.Value;
                changes.Add($"Y={y.Value}");
            }

            if (width.HasValue)
            {
                shape.Width = width.Value;
                changes.Add($"Width={width.Value}");
            }

            if (height.HasValue)
            {
                shape.Height = height.Value;
                changes.Add($"Height={height.Value}");
            }

            if (rotation.HasValue)
            {
                shape.Rotation = rotation.Value;
                changes.Add($"Rotation={rotation.Value}");
            }

            if (flipHorizontal.HasValue || flipVertical.HasValue)
            {
                var currentFrame = shape.Frame;
                var newFlipH = flipHorizontal.HasValue
                    ? flipHorizontal.Value ? NullableBool.True : NullableBool.False
                    : currentFrame.FlipH;
                var newFlipV = flipVertical.HasValue
                    ? flipVertical.Value ? NullableBool.True : NullableBool.False
                    : currentFrame.FlipV;

                shape.Frame = new ShapeFrame(
                    shape.X, shape.Y, shape.Width, shape.Height,
                    newFlipH, newFlipV, shape.Rotation);

                if (flipHorizontal.HasValue) changes.Add($"FlipH={flipHorizontal.Value}");
                if (flipVertical.HasValue) changes.Add($"FlipV={flipVertical.Value}");
            }

            if (text != null && shape is IAutoShape { TextFrame: not null } autoShape)
            {
                autoShape.TextFrame.Text = text;
                changes.Add("Text updated");
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return
                $"Shape {shapeIndex} on slide {slideIndex} edited ({string.Join(", ", changes)}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets shape fill and line format
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, fillColor, lineColor, lineWidth</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is out of range</exception>
    private Task<string> SetFormatAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var fillColor = ArgumentHelper.GetStringNullable(arguments, "fillColor");
            var lineColor = ArgumentHelper.GetStringNullable(arguments, "lineColor");
            var lineWidth = ArgumentHelper.GetFloatNullable(arguments, "lineWidth");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            var changes = new List<string>();

            if (!string.IsNullOrWhiteSpace(fillColor))
            {
                var color = ColorHelper.ParseColor(fillColor);
                shape.FillFormat.FillType = FillType.Solid;
                shape.FillFormat.SolidFillColor.Color = color;
                changes.Add($"Fill={fillColor}");
            }

            if (!string.IsNullOrWhiteSpace(lineColor))
            {
                var color = ColorHelper.ParseColor(lineColor);
                shape.LineFormat.FillFormat.FillType = FillType.Solid;
                shape.LineFormat.FillFormat.SolidFillColor.Color = color;
                changes.Add($"LineColor={lineColor}");
            }

            if (lineWidth.HasValue)
            {
                shape.LineFormat.Width = lineWidth.Value;
                changes.Add($"LineWidth={lineWidth.Value}");
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape {shapeIndex} format updated ({string.Join(", ", changes)}). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Clears shape fill and/or line format
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, clearFill, clearLine</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is out of range</exception>
    private Task<string> ClearFormatAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var clearFill = ArgumentHelper.GetBool(arguments, "clearFill", false);
            var clearLine = ArgumentHelper.GetBool(arguments, "clearLine", false);

            if (!clearFill && !clearLine)
                throw new ArgumentException("At least one of clearFill or clearLine must be true");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            var changes = new List<string>();

            if (clearFill)
            {
                shape.FillFormat.FillType = FillType.NoFill;
                changes.Add("Fill cleared");
            }

            if (clearLine)
            {
                shape.LineFormat.FillFormat.FillType = FillType.NoFill;
                changes.Add("Line cleared");
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape {shapeIndex} format cleared ({string.Join(", ", changes)}). Output: {outputPath}";
        });
    }

    #endregion

    #region Advanced Operations

    /// <summary>
    ///     Groups multiple shapes together
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndices</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when less than 2 shapes provided or indices are out of range</exception>
    private Task<string> GroupShapesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndicesArray = ArgumentHelper.GetArray(arguments, "shapeIndices");

            var shapeIndices = shapeIndicesArray
                .Select(s => s?.GetValue<int>())
                .Where(s => s.HasValue)
                .Select(s => s!.Value)
                .OrderByDescending(s => s)
                .ToList();

            if (shapeIndices.Count < 2)
                throw new ArgumentException("At least 2 shapes are required for grouping");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var shapesToGroup = new List<(IShape shape, int originalIndex)>();
            foreach (var idx in shapeIndices)
            {
                PowerPointHelper.ValidateCollectionIndex(idx, slide.Shapes.Count, "shapeIndex");
                shapesToGroup.Add((slide.Shapes[idx], idx));
            }

            // Calculate bounding box
            var minX = shapesToGroup.Min(s => s.shape.X);
            var minY = shapesToGroup.Min(s => s.shape.Y);
            var maxX = shapesToGroup.Max(s => s.shape.X + s.shape.Width);
            var maxY = shapesToGroup.Max(s => s.shape.Y + s.shape.Height);

            var groupShape = slide.Shapes.AddGroupShape();
            groupShape.X = minX;
            groupShape.Y = minY;
            groupShape.Width = maxX - minX;
            groupShape.Height = maxY - minY;

            foreach (var (shape, originalIndex) in shapesToGroup)
            {
                var clonedShape = groupShape.Shapes.AddClone(shape);
                clonedShape.X = shape.X - minX;
                clonedShape.Y = shape.Y - minY;
                slide.Shapes.RemoveAt(originalIndex);
            }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Grouped {shapeIndices.Count} shapes on slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Ungroups a group shape
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not a group</exception>
    private Task<string> UngroupShapesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            if (shape is not IGroupShape groupShape)
                throw new ArgumentException($"Shape at index {shapeIndex} is not a group");

            var groupIndex = slide.Shapes.IndexOf(groupShape);
            var shapesInGroup = groupShape.Shapes.ToList();

            var insertIndex = groupIndex;
            foreach (var s in shapesInGroup)
            {
                slide.Shapes.InsertClone(insertIndex, s);
                insertIndex++;
            }

            slide.Shapes.Remove(groupShape);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Ungrouped {shapesInGroup.Count} shapes on slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Copies a shape to another slide
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing fromSlide, toSlide, shapeIndex</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when slide or shape indices are out of range</exception>
    private Task<string> CopyShapeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var fromSlide = ArgumentHelper.GetInt(arguments, "fromSlide");
            var toSlide = ArgumentHelper.GetInt(arguments, "toSlide");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);

            PowerPointHelper.ValidateCollectionIndex(fromSlide, presentation.Slides.Count, "fromSlide");
            PowerPointHelper.ValidateCollectionIndex(toSlide, presentation.Slides.Count, "toSlide");

            var sourceSlide = presentation.Slides[fromSlide];
            PowerPointHelper.ValidateCollectionIndex(shapeIndex, sourceSlide.Shapes.Count, "shapeIndex");

            var targetSlide = presentation.Slides[toSlide];
            targetSlide.Shapes.AddClone(sourceSlide.Shapes[shapeIndex]);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape {shapeIndex} copied from slide {fromSlide} to slide {toSlide}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Reorders a shape's Z-order position
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, toIndex</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range</exception>
    private Task<string> ReorderShapeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var toIndex = ArgumentHelper.GetInt(arguments, "toIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            PowerPointHelper.ValidateCollectionIndex(shapeIndex, slide.Shapes.Count, "shapeIndex");
            PowerPointHelper.ValidateCollectionIndex(toIndex, slide.Shapes.Count, "toIndex");

            var shape = slide.Shapes[shapeIndex];
            slide.Shapes.Reorder(toIndex, shape);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Shape Z-order changed: {shapeIndex} -> {toIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Aligns multiple shapes
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndices, align, alignToSlide</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when less than 2 shapes or invalid alignment</exception>
    private Task<string> AlignShapesAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var alignStr = ArgumentHelper.GetString(arguments, "align");
            var shapeIndicesArray = ArgumentHelper.GetArray(arguments, "shapeIndices");
            var alignToSlide = ArgumentHelper.GetBool(arguments, "alignToSlide", false);

            var shapeIndices = shapeIndicesArray.Select(x => x?.GetValue<int>() ?? -1).ToArray();

            if (shapeIndices.Length < 2)
                throw new ArgumentException("At least 2 shapes are required for alignment");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            foreach (var idx in shapeIndices)
                PowerPointHelper.ValidateCollectionIndex(idx, slide.Shapes.Count, "shapeIndex");

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
                    case "left": s.X = refBox.X; break;
                    case "center": s.X = refBox.X + (refBox.W - s.Width) / 2f; break;
                    case "right": s.X = refBox.X + refBox.W - s.Width; break;
                    case "top": s.Y = refBox.Y; break;
                    case "middle": s.Y = refBox.Y + (refBox.H - s.Height) / 2f; break;
                    case "bottom": s.Y = refBox.Y + refBox.H - s.Height; break;
                    default:
                        throw new ArgumentException("align must be: left|center|right|top|middle|bottom");
                }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Aligned {shapeIndices.Length} shapes: {alignStr}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Flips a shape horizontally or vertically
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing slideIndex, shapeIndex, flipHorizontal, flipVertical</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when no flip direction specified</exception>
    private Task<string> FlipShapeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var flipHorizontal = ArgumentHelper.GetBoolNullable(arguments, "flipHorizontal");
            var flipVertical = ArgumentHelper.GetBoolNullable(arguments, "flipVertical");

            if (!flipHorizontal.HasValue && !flipVertical.HasValue)
                throw new ArgumentException("At least one of flipHorizontal or flipVertical must be provided");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);

            var currentFrame = shape.Frame;
            var newFlipH = flipHorizontal.HasValue
                ? flipHorizontal.Value ? NullableBool.True : NullableBool.False
                : currentFrame.FlipH;
            var newFlipV = flipVertical.HasValue
                ? flipVertical.Value ? NullableBool.True : NullableBool.False
                : currentFrame.FlipV;

            shape.Frame = new ShapeFrame(
                currentFrame.X, currentFrame.Y, currentFrame.Width, currentFrame.Height,
                newFlipH, newFlipV, currentFrame.Rotation);

            presentation.Save(outputPath, SaveFormat.Pptx);

            var flipDesc = new List<string>();
            if (flipHorizontal.HasValue) flipDesc.Add($"H={flipHorizontal.Value}");
            if (flipVertical.HasValue) flipDesc.Add($"V={flipVertical.Value}");

            return $"Shape {shapeIndex} flipped ({string.Join(", ", flipDesc)}). Output: {outputPath}";
        });
    }

    #endregion
}