using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint slides (add, delete, get info, move, duplicate, hide, clear, edit)
/// </summary>
[McpServerToolType]
public class PptSlideTool
{
    /// <summary>
    ///     Handler registry for slide operations.
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
    ///     Initializes a new instance of the <see cref="PptSlideTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptSlideTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Slide");
    }

    /// <summary>
    ///     Executes a PowerPoint slide operation (add, delete, get_info, move, duplicate, hide, clear, edit).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, get_info, move, duplicate, hide, clear, edit.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to input path).</param>
    /// <param name="slideIndex">Slide index (0-based, required for most operations).</param>
    /// <param name="layoutType">Slide layout type: Blank, Title, TitleOnly, TwoColumn, SectionHeader.</param>
    /// <param name="fromIndex">Source slide index (0-based, required for move operation).</param>
    /// <param name="toIndex">Target slide index (0-based, required for move operation).</param>
    /// <param name="insertAt">Target index to insert clone (0-based, optional, for duplicate, default: append).</param>
    /// <param name="slideIndices">
    ///     Slide indices array as JSON (optional, for hide operation, if not provided affects all
    ///     slides).
    /// </param>
    /// <param name="hidden">Hide slides (true) or show (false, required for hide operation).</param>
    /// <param name="layoutIndex">Layout index (0-based, optional, for edit operation).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_slide")]
    [Description(
        @"Manage PowerPoint slides. Supports 8 operations: add, delete, get_info, move, duplicate, hide, clear, edit.

Usage examples:
- Add slide: ppt_slide(operation='add', path='presentation.pptx', layoutType='Blank')
- Delete slide: ppt_slide(operation='delete', path='presentation.pptx', slideIndex=0)
- Get info: ppt_slide(operation='get_info', path='presentation.pptx')
- Move slide: ppt_slide(operation='move', path='presentation.pptx', fromIndex=0, toIndex=2)
- Duplicate slide: ppt_slide(operation='duplicate', path='presentation.pptx', slideIndex=0)
- Hide slide: ppt_slide(operation='hide', path='presentation.pptx', slideIndex=0, hidden=true)
- Clear slide: ppt_slide(operation='clear', path='presentation.pptx', slideIndex=0)
- Edit slide: ppt_slide(operation='edit', path='presentation.pptx', slideIndex=0)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a new slide (required params: path)
- 'delete': Delete a slide (required params: path, slideIndex)
- 'get_info': Get slides info (required params: path)
- 'move': Move a slide (required params: path, fromIndex, toIndex)
- 'duplicate': Duplicate a slide (required params: path, slideIndex)
- 'hide': Hide/show a slide (required params: path, slideIndex, hidden)
- 'clear': Clear slide content (required params: path, slideIndex)
- 'edit': Edit slide properties (required params: path, slideIndex)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for most operations)")]
        int? slideIndex = null,
        [Description("Slide layout type: Blank, Title, TitleOnly, TwoColumn, SectionHeader")]
        string layoutType = "Blank",
        [Description("Source slide index (0-based, required for move operation)")]
        int? fromIndex = null,
        [Description("Target slide index (0-based, required for move operation)")]
        int? toIndex = null,
        [Description("Target index to insert clone (0-based, optional, for duplicate, default: append)")]
        int? insertAt = null,
        [Description("Slide indices array as JSON (optional, for hide operation, if not provided affects all slides)")]
        string? slideIndices = null,
        [Description("Hide slides (true) or show (false, required for hide operation)")]
        bool hidden = false,
        [Description("Layout index (0-based, optional, for edit operation)")]
        int? layoutIndex = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, layoutType, fromIndex, toIndex, insertAt,
            slideIndices, hidden, layoutIndex);

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

        if (operation.ToLowerInvariant() == "get_info")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int? slideIndex,
        string layoutType,
        int? fromIndex,
        int? toIndex,
        int? insertAt,
        string? slideIndices,
        bool hidden,
        int? layoutIndex)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "add":
                parameters.Set("layoutType", layoutType);
                break;

            case "delete":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                break;

            case "get_info":
                break;

            case "move":
                if (fromIndex.HasValue) parameters.Set("fromIndex", fromIndex.Value);
                if (toIndex.HasValue) parameters.Set("toIndex", toIndex.Value);
                break;

            case "duplicate":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                if (insertAt.HasValue) parameters.Set("insertAt", insertAt.Value);
                break;

            case "hide":
                if (slideIndices != null) parameters.Set("slideIndices", slideIndices);
                parameters.Set("hidden", hidden);
                break;

            case "clear":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                break;

            case "edit":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                if (layoutIndex.HasValue) parameters.Set("layoutIndex", layoutIndex.Value);
                break;
        }

        return parameters;
    }
}
