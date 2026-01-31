using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint text formatting (batch format text)
///     Merges: PptBatchFormatTextTool
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.TextFormat")]
[McpServerToolType]
public class PptTextFormatTool
{
    /// <summary>
    ///     Handler registry for text format operations.
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
    ///     Initializes a new instance of the <see cref="PptTextFormatTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptTextFormatTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.TextFormat");
    }

    /// <summary>
    ///     Executes a PowerPoint text format operation (batch format text).
    /// </summary>
    /// <param name="operation">The operation to perform: format.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to input path).</param>
    /// <param name="slideIndices">Slide indices to apply as JSON array (optional; default all).</param>
    /// <param name="shapeIndices">Shape indices to apply as JSON array (optional; default all shapes on targeted slides).</param>
    /// <param name="fontName">Font name (optional).</param>
    /// <param name="fontSize">Font size (optional).</param>
    /// <param name="bold">Bold (optional).</param>
    /// <param name="italic">Italic (optional).</param>
    /// <param name="color">Text color: Hex (#FF5500, #RGB) or named color (Red, Blue, DarkGreen) (optional).</param>
    /// <param name="alignment">Text alignment: Left, Center, Right, Justify, Distributed (optional).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slide or shape index is out of range.</exception>
    [McpServerTool(
        Name = "ppt_text_format",
        Title = "PowerPoint Text Format Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Batch format PowerPoint text. Formats font, size, bold, italic, color, alignment across slides.
Applies to text in AutoShapes and Table cells.

Color format: Hex color code (e.g., #FF5500, #RGB, #RRGGBB) or named colors (e.g., Red, Blue, DarkGreen).
Alignment values: Left, Center, Right, Justify, Distributed.

Usage examples:
- Format all slides: ppt_text_format(operation='format', path='presentation.pptx', fontName='Arial', fontSize=14, bold=true)
- Format specific slides: ppt_text_format(operation='format', path='presentation.pptx', slideIndices='[0,1,2]', fontName='Times New Roman', fontSize=12)
- Format specific shapes: ppt_text_format(operation='format', path='presentation.pptx', slideIndices='[0]', shapeIndices='[1,2]', fontSize=18)
- Format with color: ppt_text_format(operation='format', path='presentation.pptx', color='#FF0000') or ppt_text_format(operation='format', path='presentation.pptx', color='Red')
- Center align text: ppt_text_format(operation='format', path='presentation.pptx', alignment='Center')")]
    public object Execute(
        [Description("Operation: format")] string operation = "format",
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide indices to apply as JSON array (optional; default all)")]
        string? slideIndices = null,
        [Description(
            "Shape indices (0-based) to apply as JSON array (optional; default all shapes on targeted slides)")]
        string? shapeIndices = null,
        [Description("Font name (optional)")] string? fontName = null,
        [Description("Font size (optional)")] double? fontSize = null,
        [Description("Bold (optional)")] bool? bold = null,
        [Description("Italic (optional)")] bool? italic = null,
        [Description("Text color: Hex (#FF5500, #RGB) or named color (Red, Blue, DarkGreen) (optional)")]
        string? color = null,
        [Description("Text alignment: Left, Center, Right, Justify, Distributed (optional)")]
        string? alignment = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters =
            BuildParameters(slideIndices, shapeIndices, fontName, fontSize, bold, italic, color, alignment);

        var handler = _handlerRegistry.GetHandler(operation);

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

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    /// <param name="slideIndices">Slide indices to apply as JSON array.</param>
    /// <param name="shapeIndices">Shape indices to apply as JSON array.</param>
    /// <param name="fontName">The font name to apply.</param>
    /// <param name="fontSize">The font size to apply.</param>
    /// <param name="bold">Whether to apply bold formatting.</param>
    /// <param name="italic">Whether to apply italic formatting.</param>
    /// <param name="color">The text color (hex or named color).</param>
    /// <param name="alignment">The text alignment (Left, Center, Right, Justify, Distributed).</param>
    /// <returns>OperationParameters configured for text formatting.</returns>
    private static OperationParameters BuildParameters(
        string? slideIndices,
        string? shapeIndices,
        string? fontName,
        double? fontSize,
        bool? bold,
        bool? italic,
        string? color,
        string? alignment)
    {
        var parameters = new OperationParameters();

        if (slideIndices != null) parameters.Set("slideIndices", slideIndices);
        if (shapeIndices != null) parameters.Set("shapeIndices", shapeIndices);
        if (fontName != null) parameters.Set("fontName", fontName);
        if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
        if (bold.HasValue) parameters.Set("bold", bold.Value);
        if (italic.HasValue) parameters.Set("italic", italic.Value);
        if (color != null) parameters.Set("color", color);
        if (alignment != null) parameters.Set("alignment", alignment);

        return parameters;
    }
}
