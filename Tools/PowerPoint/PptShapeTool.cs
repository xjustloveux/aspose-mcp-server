using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
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
    ///     Handler registry for shape operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

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
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Shape");
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

        var handlerOperation = MapOperationToHandler(operation);
        var parameters = BuildParameters(operation, slideIndex, shapeIndex, shapeIndices, x, y, width, height,
            rotation, text, fillColor, lineColor, lineWidth, clearFill, clearLine, fromSlide, toSlide, toIndex,
            align, alignToSlide, flipHorizontal, flipVertical);

        var handler = _handlerRegistry.GetHandler(handlerOperation);

        var operationContext = new OperationContext<Presentation>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        var op = operation.ToLowerInvariant();
        if (op == "get" || op == "get_details")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Maps Tool operation name to Handler operation name.
    /// </summary>
    private static string MapOperationToHandler(string operation)
    {
        return operation.ToLowerInvariant() switch
        {
            "get" => "get_shapes",
            "get_details" => "get_shape_details",
            _ => operation
        };
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int? slideIndex,
        int? shapeIndex,
        int[]? shapeIndices,
        float? x,
        float? y,
        float? width,
        float? height,
        float? rotation,
        string? text,
        string? fillColor,
        string? lineColor,
        float? lineWidth,
        bool clearFill,
        bool clearLine,
        int? fromSlide,
        int? toSlide,
        int? toIndex,
        string? align,
        bool alignToSlide,
        bool? flipHorizontal,
        bool? flipVertical)
    {
        var parameters = new OperationParameters();

        return operation.ToLowerInvariant() switch
        {
            "get" or "get_details" or "delete" or "ungroup"
                => BuildSlideShapeParameters(parameters, slideIndex, shapeIndex),
            "edit" => BuildEditParameters(parameters, slideIndex, shapeIndex, x, y, width, height, rotation, text),
            "set_format" => BuildSetFormatParameters(parameters, slideIndex, shapeIndex, fillColor, lineColor,
                lineWidth),
            "clear_format" => BuildClearFormatParameters(parameters, slideIndex, shapeIndex, clearFill, clearLine),
            "group" or "align" => BuildGroupAlignParameters(parameters, slideIndex, shapeIndices, align, alignToSlide),
            "copy" => BuildCopyParameters(parameters, fromSlide, toSlide, shapeIndex),
            "reorder" => BuildReorderParameters(parameters, slideIndex, shapeIndex, toIndex),
            "flip" => BuildFlipParameters(parameters, slideIndex, shapeIndex, flipHorizontal, flipVertical),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for basic slide/shape operations (get, get_details, delete, ungroup).
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <returns>OperationParameters configured for basic slide/shape operations.</returns>
    private static OperationParameters BuildSlideShapeParameters(OperationParameters parameters, int? slideIndex,
        int? shapeIndex)
    {
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit shape operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <param name="rotation">The rotation in degrees.</param>
    /// <param name="text">The text content for AutoShape.</param>
    /// <returns>OperationParameters configured for the edit operation.</returns>
    private static OperationParameters BuildEditParameters(OperationParameters parameters, int? slideIndex,
        int? shapeIndex, float? x, float? y, float? width, float? height, float? rotation, string? text)
    {
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (x.HasValue) parameters.Set("x", x.Value);
        if (y.HasValue) parameters.Set("y", y.Value);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        if (rotation.HasValue) parameters.Set("rotation", rotation.Value);
        if (text != null) parameters.Set("text", text);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set format operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="fillColor">The fill color in hex format (e.g., '#FF0000').</param>
    /// <param name="lineColor">The line color in hex format.</param>
    /// <param name="lineWidth">The line width in points.</param>
    /// <returns>OperationParameters configured for the set format operation.</returns>
    private static OperationParameters BuildSetFormatParameters(OperationParameters parameters, int? slideIndex,
        int? shapeIndex, string? fillColor, string? lineColor, float? lineWidth)
    {
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (fillColor != null) parameters.Set("fillColor", fillColor);
        if (lineColor != null) parameters.Set("lineColor", lineColor);
        if (lineWidth.HasValue) parameters.Set("lineWidth", lineWidth.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the clear format operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="clearFill">Whether to clear the fill.</param>
    /// <param name="clearLine">Whether to clear the line.</param>
    /// <returns>OperationParameters configured for the clear format operation.</returns>
    private static OperationParameters BuildClearFormatParameters(OperationParameters parameters, int? slideIndex,
        int? shapeIndex, bool clearFill, bool clearLine)
    {
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        parameters.Set("clearFill", clearFill);
        parameters.Set("clearLine", clearLine);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the group and align operations.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndices">The array of shape indices to group or align.</param>
    /// <param name="align">The alignment type: left, center, right, top, middle, bottom.</param>
    /// <param name="alignToSlide">Whether to align to slide bounds.</param>
    /// <returns>OperationParameters configured for group/align operations.</returns>
    private static OperationParameters BuildGroupAlignParameters(OperationParameters parameters, int? slideIndex,
        int[]? shapeIndices, string? align, bool alignToSlide)
    {
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndices != null) parameters.Set("shapeIndices", shapeIndices);
        if (align != null) parameters.Set("align", align);
        parameters.Set("alignToSlide", alignToSlide);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the copy shape operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="fromSlide">The source slide index (0-based).</param>
    /// <param name="toSlide">The target slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index to copy (0-based).</param>
    /// <returns>OperationParameters configured for the copy operation.</returns>
    private static OperationParameters BuildCopyParameters(OperationParameters parameters, int? fromSlide,
        int? toSlide, int? shapeIndex)
    {
        if (fromSlide.HasValue) parameters.Set("fromSlide", fromSlide.Value);
        if (toSlide.HasValue) parameters.Set("toSlide", toSlide.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the reorder shape operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="toIndex">The target Z-order index.</param>
    /// <returns>OperationParameters configured for the reorder operation.</returns>
    private static OperationParameters BuildReorderParameters(OperationParameters parameters, int? slideIndex,
        int? shapeIndex, int? toIndex)
    {
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (toIndex.HasValue) parameters.Set("toIndex", toIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the flip shape operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="flipHorizontal">Whether to flip horizontally.</param>
    /// <param name="flipVertical">Whether to flip vertically.</param>
    /// <returns>OperationParameters configured for the flip operation.</returns>
    private static OperationParameters BuildFlipParameters(OperationParameters parameters, int? slideIndex,
        int? shapeIndex, bool? flipHorizontal, bool? flipVertical)
    {
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (flipHorizontal.HasValue) parameters.Set("flipHorizontal", flipHorizontal.Value);
        if (flipVertical.HasValue) parameters.Set("flipVertical", flipVertical.Value);
        return parameters;
    }
}
