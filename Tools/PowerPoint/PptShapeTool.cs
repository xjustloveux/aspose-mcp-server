using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Core.ShapeDetailProviders;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint shapes.
///     Supports 12 operations: get, get_details, delete, edit, set_format, clear_format,
///     group, ungroup, copy, reorder, align, flip
/// </summary>
[McpServerToolType]
public class PptShapeTool
{
    /// <summary>
    ///     JSON serializer options for consistent output formatting with support for floating point literals.
    /// </summary>
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        NumberHandling = JsonNumberHandling.AllowNamedFloatingPointLiterals
    };

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptShapeTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptShapeTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PowerPoint shape operation (get, get_details, delete, edit, set_format, clear_format, group, ungroup,
    ///     copy, reorder, align, flip).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: get, get_details, delete, edit, set_format, clear_format, group,
    ///     ungroup, copy, reorder, align, flip.
    /// </param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="slideIndex">Slide index (0-based, required except for copy).</param>
    /// <param name="shapeIndex">Shape index (0-based, for single-shape operations).</param>
    /// <param name="shapeIndices">Shape indices array (for group/align).</param>
    /// <param name="x">X position in points (for edit).</param>
    /// <param name="y">Y position in points (for edit).</param>
    /// <param name="width">Width in points (for edit).</param>
    /// <param name="height">Height in points (for edit).</param>
    /// <param name="rotation">Rotation in degrees (for edit).</param>
    /// <param name="text">Text content for AutoShape (for edit).</param>
    /// <param name="fillColor">Fill color hex, e.g. #FF0000 (for set_format).</param>
    /// <param name="lineColor">Line color hex (for set_format).</param>
    /// <param name="lineWidth">Line width in points (for set_format).</param>
    /// <param name="clearFill">Clear fill (for clear_format).</param>
    /// <param name="clearLine">Clear line (for clear_format).</param>
    /// <param name="fromSlide">Source slide index (for copy).</param>
    /// <param name="toSlide">Target slide index (for copy).</param>
    /// <param name="toIndex">Target Z-order index (for reorder).</param>
    /// <param name="align">Alignment: left|center|right|top|middle|bottom (for align).</param>
    /// <param name="alignToSlide">Align to slide bounds (for align, default: false).</param>
    /// <param name="flipHorizontal">Flip horizontally (for flip/edit).</param>
    /// <param name="flipVertical">Flip vertically (for flip/edit).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_shape")]
    [Description(@"Unified PowerPoint shape management tool. Supports 12 operations.

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
- Flip: ppt_shape(operation='flip', path='file.pptx', slideIndex=0, shapeIndex=0, flipHorizontal=true)")]
    public string Execute(
        [Description(
            "Operation: get, get_details, delete, edit, set_format, clear_format, group, ungroup, copy, reorder, align, flip")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required except for copy)")]
        int? slideIndex = null,
        [Description("Shape index (0-based, for single-shape operations)")]
        int? shapeIndex = null,
        [Description("Shape indices array (for group/align)")]
        int[]? shapeIndices = null,
        [Description("X position in points (for edit)")]
        float? x = null,
        [Description("Y position in points (for edit)")]
        float? y = null,
        [Description("Width in points (for edit)")]
        float? width = null,
        [Description("Height in points (for edit)")]
        float? height = null,
        [Description("Rotation in degrees (for edit)")]
        float? rotation = null,
        [Description("Text content for AutoShape (for edit)")]
        string? text = null,
        [Description("Fill color hex, e.g. #FF0000 (for set_format)")]
        string? fillColor = null,
        [Description("Line color hex (for set_format)")]
        string? lineColor = null,
        [Description("Line width in points (for set_format)")]
        float? lineWidth = null,
        [Description("Clear fill (for clear_format)")]
        bool clearFill = false,
        [Description("Clear line (for clear_format)")]
        bool clearLine = false,
        [Description("Source slide index (for copy)")]
        int? fromSlide = null,
        [Description("Target slide index (for copy)")]
        int? toSlide = null,
        [Description("Target Z-order index (for reorder)")]
        int? toIndex = null,
        [Description("Alignment: left|center|right|top|middle|bottom (for align)")]
        string? align = null,
        [Description("Align to slide bounds (for align, default: false)")]
        bool alignToSlide = false,
        [Description("Flip horizontally (for flip/edit)")]
        bool? flipHorizontal = null,
        [Description("Flip vertically (for flip/edit)")]
        bool? flipVertical = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "get" => GetShapes(ctx, slideIndex),
            "get_details" => GetShapeDetails(ctx, slideIndex, shapeIndex),
            "delete" => DeleteShape(ctx, outputPath, slideIndex, shapeIndex),
            "edit" => EditShape(ctx, outputPath, slideIndex, shapeIndex, x, y, width, height, rotation, flipHorizontal,
                flipVertical, text),
            "set_format" => SetFormat(ctx, outputPath, slideIndex, shapeIndex, fillColor, lineColor, lineWidth),
            "clear_format" => ClearFormat(ctx, outputPath, slideIndex, shapeIndex, clearFill, clearLine),
            "group" => GroupShapes(ctx, outputPath, slideIndex, shapeIndices),
            "ungroup" => UngroupShapes(ctx, outputPath, slideIndex, shapeIndex),
            "copy" => CopyShape(ctx, outputPath, fromSlide, toSlide, shapeIndex),
            "reorder" => ReorderShape(ctx, outputPath, slideIndex, shapeIndex, toIndex),
            "align" => AlignShapes(ctx, outputPath, slideIndex, shapeIndices, align, alignToSlide),
            "flip" => FlipShape(ctx, outputPath, slideIndex, shapeIndex, flipHorizontal, flipVertical),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Gets all shapes from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <returns>A JSON string containing shape information including types, positions, and sizes.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is not provided.</exception>
    private static string GetShapes(DocumentContext<Presentation> ctx, int? slideIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for get operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);

        List<object> shapesList = [];
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
    }

    /// <summary>
    ///     Gets detailed information about a specific shape.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndex">The zero-based index of the shape.</param>
    /// <returns>
    ///     A JSON string containing detailed shape information including position, size, rotation, fill, line, and
    ///     type-specific properties.
    /// </returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is not provided.</exception>
    private static string GetShapeDetails(DocumentContext<Presentation> ctx, int? slideIndex, int? shapeIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for get_details operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for get_details operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

        var frame = shape.Frame;
        var (typeName, typeSpecificProperties) = ShapeDetailProviderFactory.GetShapeDetails(shape, presentation);

        string? fillColorHex = null;
        if (shape.FillFormat.FillType == FillType.Solid)
        {
            var color = shape.FillFormat.SolidFillColor.Color;
            fillColorHex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

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
    }

    /// <summary>
    ///     Deletes a shape from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndex">The zero-based index of the shape to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is not provided or out of range.</exception>
    private static string DeleteShape(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? shapeIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for delete operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        PowerPointHelper.ValidateCollectionIndex(shapeIndex.Value, slide.Shapes.Count, "shapeIndex");

        slide.Shapes.RemoveAt(shapeIndex.Value);

        ctx.Save(outputPath);
        var remaining = slide.Shapes.Count;

        var result = $"Shape {shapeIndex} deleted from slide {slideIndex} ({remaining} remaining).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Edits shape properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndex">The zero-based index of the shape.</param>
    /// <param name="x">The X position in points (optional).</param>
    /// <param name="y">The Y position in points (optional).</param>
    /// <param name="width">The width in points (optional).</param>
    /// <param name="height">The height in points (optional).</param>
    /// <param name="rotation">The rotation in degrees (optional).</param>
    /// <param name="flipHorizontal">True to flip horizontally (optional).</param>
    /// <param name="flipVertical">True to flip vertically (optional).</param>
    /// <param name="text">The text content for AutoShape (optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is not provided.</exception>
    private static string EditShape(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? shapeIndex,
        float? x, float? y, float? width, float? height, float? rotation, bool? flipHorizontal, bool? flipVertical,
        string? text)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for edit operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for edit operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

        List<string> changes = [];

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

        ctx.Save(outputPath);

        var result = $"Shape {shapeIndex} on slide {slideIndex} edited ({string.Join(", ", changes)}).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets shape fill and line format.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndex">The zero-based index of the shape.</param>
    /// <param name="fillColor">The fill color in hex format (e.g., #FF0000).</param>
    /// <param name="lineColor">The line color in hex format.</param>
    /// <param name="lineWidth">The line width in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is not provided.</exception>
    private static string SetFormat(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? shapeIndex,
        string? fillColor, string? lineColor, float? lineWidth)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for set_format operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for set_format operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

        List<string> changes = [];

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

        ctx.Save(outputPath);

        var result = $"Shape {shapeIndex} format updated ({string.Join(", ", changes)}).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Clears shape fill and/or line format.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndex">The zero-based index of the shape.</param>
    /// <param name="clearFill">True to clear the fill format.</param>
    /// <param name="clearLine">True to clear the line format.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex or shapeIndex is not provided, or neither clearFill nor
    ///     clearLine is true.
    /// </exception>
    private static string ClearFormat(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? shapeIndex,
        bool clearFill, bool clearLine)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for clear_format operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for clear_format operation");
        if (!clearFill && !clearLine)
            throw new ArgumentException("At least one of clearFill or clearLine must be true");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

        List<string> changes = [];

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

        ctx.Save(outputPath);

        var result = $"Shape {shapeIndex} format cleared ({string.Join(", ", changes)}).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Groups multiple shapes together.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndices">Array of shape indices to group.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex is not provided or less than 2 shapes are specified.</exception>
    private static string GroupShapes(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int[]? shapeIndices)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for group operation");
        if (shapeIndices == null || shapeIndices.Length < 2)
            throw new ArgumentException("At least 2 shapes are required for grouping");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);

        var sortedIndices = shapeIndices.OrderByDescending(s => s).ToList();

        List<(IShape shape, int originalIndex)> shapesToGroup = [];
        foreach (var idx in sortedIndices)
        {
            PowerPointHelper.ValidateCollectionIndex(idx, slide.Shapes.Count, "shapeIndex");
            shapesToGroup.Add((slide.Shapes[idx], idx));
        }

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

        ctx.Save(outputPath);

        var result = $"Grouped {shapeIndices.Length} shapes on slide {slideIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Ungroups a group shape.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndex">The zero-based index of the group shape to ungroup.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or shapeIndex is not provided, or the shape is not a group.</exception>
    private static string UngroupShapes(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? shapeIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for ungroup operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for ungroup operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

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

        ctx.Save(outputPath);

        var result = $"Ungrouped {shapesInGroup.Count} shapes on slide {slideIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Copies a shape to another slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="fromSlide">The source slide index (0-based).</param>
    /// <param name="toSlide">The target slide index (0-based).</param>
    /// <param name="shapeIndex">The zero-based index of the shape to copy.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when fromSlide, toSlide, or shapeIndex is not provided or out of range.</exception>
    private static string CopyShape(DocumentContext<Presentation> ctx, string? outputPath, int? fromSlide, int? toSlide,
        int? shapeIndex)
    {
        if (!fromSlide.HasValue)
            throw new ArgumentException("fromSlide is required for copy operation");
        if (!toSlide.HasValue)
            throw new ArgumentException("toSlide is required for copy operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for copy operation");

        var presentation = ctx.Document;

        PowerPointHelper.ValidateCollectionIndex(fromSlide.Value, presentation.Slides.Count, "fromSlide");
        PowerPointHelper.ValidateCollectionIndex(toSlide.Value, presentation.Slides.Count, "toSlide");

        var sourceSlide = presentation.Slides[fromSlide.Value];
        PowerPointHelper.ValidateCollectionIndex(shapeIndex.Value, sourceSlide.Shapes.Count, "shapeIndex");

        var targetSlide = presentation.Slides[toSlide.Value];
        targetSlide.Shapes.AddClone(sourceSlide.Shapes[shapeIndex.Value]);

        ctx.Save(outputPath);

        var result = $"Shape {shapeIndex} copied from slide {fromSlide} to slide {toSlide}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Reorders a shape's Z-order position.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndex">The current zero-based index of the shape.</param>
    /// <param name="toIndex">The target Z-order index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex, shapeIndex, or toIndex is not provided or out of range.</exception>
    private static string ReorderShape(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? shapeIndex, int? toIndex)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for reorder operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for reorder operation");
        if (!toIndex.HasValue)
            throw new ArgumentException("toIndex is required for reorder operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);

        PowerPointHelper.ValidateCollectionIndex(shapeIndex.Value, slide.Shapes.Count, "shapeIndex");
        PowerPointHelper.ValidateCollectionIndex(toIndex.Value, slide.Shapes.Count, "toIndex");

        var shape = slide.Shapes[shapeIndex.Value];
        slide.Shapes.Reorder(toIndex.Value, shape);

        ctx.Save(outputPath);

        var result = $"Shape Z-order changed: {shapeIndex} -> {toIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Aligns multiple shapes.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndices">Array of shape indices to align.</param>
    /// <param name="align">The alignment type: left, center, right, top, middle, or bottom.</param>
    /// <param name="alignToSlide">True to align to slide bounds, false to align to shape bounds.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex or align is not provided, less than 2 shapes are specified,
    ///     or align value is invalid.
    /// </exception>
    private static string AlignShapes(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int[]? shapeIndices, string? align, bool alignToSlide)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for align operation");
        if (string.IsNullOrEmpty(align))
            throw new ArgumentException("align is required for align operation");
        if (shapeIndices == null || shapeIndices.Length < 2)
            throw new ArgumentException("At least 2 shapes are required for alignment");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);

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
            switch (align.ToLower())
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

        ctx.Save(outputPath);

        var result = $"Aligned {shapeIndices.Length} shapes: {align}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Flips a shape horizontally or vertically.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndex">The zero-based index of the shape.</param>
    /// <param name="flipHorizontal">True to flip horizontally (optional).</param>
    /// <param name="flipVertical">True to flip vertically (optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex or shapeIndex is not provided, or neither flip option is
    ///     specified.
    /// </exception>
    private static string FlipShape(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? shapeIndex, bool? flipHorizontal, bool? flipVertical)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for flip operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for flip operation");
        if (!flipHorizontal.HasValue && !flipVertical.HasValue)
            throw new ArgumentException("At least one of flipHorizontal or flipVertical must be provided");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

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

        ctx.Save(outputPath);

        List<string> flipDesc = [];
        if (flipHorizontal.HasValue) flipDesc.Add($"H={flipHorizontal.Value}");
        if (flipVertical.HasValue) flipDesc.Add($"V={flipVertical.Value}");

        var result = $"Shape {shapeIndex} flipped ({string.Join(", ", flipDesc)}).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }
}