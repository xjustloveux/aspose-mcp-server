using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint text (add, edit, replace)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Text")]
[McpServerToolType]
public class PptTextTool
{
    /// <summary>
    ///     Handler registry for text operations.
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
    ///     Initializes a new instance of the <see cref="PptTextTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptTextTool(DocumentSessionManager? sessionManager = null, ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Text");
    }

    /// <summary>
    ///     Executes a PowerPoint text operation (add, edit, replace).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, replace.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to input path).</param>
    /// <param name="slideIndex">Slide index (0-based, required for add/edit).</param>
    /// <param name="shapeIndex">Shape index (0-based, required for edit).</param>
    /// <param name="text">Text content (required for add/edit).</param>
    /// <param name="findText">Text to find (required for replace).</param>
    /// <param name="replaceText">Text to replace with (required for replace).</param>
    /// <param name="matchCase">Match case (optional, for replace, default: false).</param>
    /// <param name="x">X position in points (optional, for add, default: 50).</param>
    /// <param name="y">Y position in points (optional, for add, default: 50).</param>
    /// <param name="width">Text box width in points (optional, for add, default: 400).</param>
    /// <param name="height">Text box height in points (optional, for add, default: 100).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "ppt_text",
        Title = "PowerPoint Text Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage PowerPoint text. Supports 3 operations: add, edit, replace.
Searches text in AutoShapes, GroupShapes (recursive), and Table cells.

Coordinate unit: 1 inch = 72 points.

Usage examples:
- Add text: ppt_text(operation='add', path='presentation.pptx', slideIndex=0, text='Hello World', x=100, y=100)
- Edit text: ppt_text(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, text='Updated Text')
- Replace text: ppt_text(operation='replace', path='presentation.pptx', findText='old', replaceText='new')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add': Add text to slide (required params: path, slideIndex, text)
- 'edit': Edit text in shape (required params: path, slideIndex, shapeIndex, text)
- 'replace': Replace text in presentation (required params: path, findText, replaceText)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for add/edit)")]
        int? slideIndex = null,
        [Description("Shape index (0-based, required for edit)")]
        int? shapeIndex = null,
        [Description("Text content (required for add/edit)")]
        string? text = null,
        [Description("Text to find (required for replace)")]
        string? findText = null,
        [Description("Text to replace with (required for replace)")]
        string? replaceText = null,
        [Description("Match case (optional, for replace, default: false)")]
        bool matchCase = false,
        [Description("X position in points (optional, for add, default: 50)")]
        float x = 50,
        [Description("Y position in points (optional, for add, default: 50)")]
        float y = 50,
        [Description("Text box width in points (optional, for add, default: 400)")]
        float width = 400,
        [Description("Text box height in points (optional, for add, default: 100)")]
        float height = 100)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, shapeIndex, text, findText, replaceText,
            matchCase, x, y, width, height);

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
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int? slideIndex,
        int? shapeIndex,
        string? text,
        string? findText,
        string? replaceText,
        bool matchCase,
        float x,
        float y,
        float width,
        float height)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(slideIndex, text, x, y, width, height),
            "edit" => BuildEditParameters(slideIndex, shapeIndex, text),
            "replace" => BuildReplaceParameters(findText, replaceText, matchCase),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add text operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="text">The text content to add.</param>
    /// <param name="x">X position in points.</param>
    /// <param name="y">Y position in points.</param>
    /// <param name="width">Text box width in points.</param>
    /// <param name="height">Text box height in points.</param>
    /// <returns>OperationParameters configured for adding text.</returns>
    private static OperationParameters BuildAddParameters(int? slideIndex, string? text, float x, float y, float width,
        float height)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (text != null) parameters.Set("text", text);
        parameters.Set("x", x);
        parameters.Set("y", y);
        parameters.Set("width", width);
        parameters.Set("height", height);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit text operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="text">The new text content.</param>
    /// <returns>OperationParameters configured for editing text.</returns>
    private static OperationParameters BuildEditParameters(int? slideIndex, int? shapeIndex, string? text)
    {
        var parameters = new OperationParameters();
        if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (text != null) parameters.Set("text", text);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the replace text operation.
    /// </summary>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="matchCase">Whether to match case.</param>
    /// <returns>OperationParameters configured for replacing text.</returns>
    private static OperationParameters BuildReplaceParameters(string? findText, string? replaceText, bool matchCase)
    {
        var parameters = new OperationParameters();
        if (findText != null) parameters.Set("findText", findText);
        if (replaceText != null) parameters.Set("replaceText", replaceText);
        parameters.Set("matchCase", matchCase);
        return parameters;
    }
}
