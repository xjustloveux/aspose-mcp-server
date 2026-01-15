using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint layouts.
///     Supports: set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme
/// </summary>
[McpServerToolType]
public class PptLayoutTool
{
    /// <summary>
    ///     Handler registry for layout operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptLayoutTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptLayoutTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Layout");
    }

    /// <summary>
    ///     Executes a PowerPoint layout operation (set, get_layouts, get_masters, apply_master, apply_layout_range,
    ///     apply_theme).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: set, get_layouts, get_masters, apply_master, apply_layout_range,
    ///     apply_theme.
    /// </param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="slideIndex">Slide index (0-based, required for set).</param>
    /// <param name="layout">
    ///     Layout type: Title, TitleOnly, Blank, TwoColumn, SectionHeader, TitleAndContent, ObjectAndText,
    ///     PictureAndCaption.
    /// </param>
    /// <param name="masterIndex">Master index (0-based, optional for get_layouts, required for apply_master).</param>
    /// <param name="layoutIndex">Layout index under master (0-based, required for apply_master).</param>
    /// <param name="slideIndices">Slide indices array as JSON (required for apply_layout_range, optional for apply_master).</param>
    /// <param name="themePath">Theme template file path (.potx/.pptx, required for apply_theme).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_layout")]
    [Description(
        @"Manage PowerPoint layouts. Supports 6 operations: set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme.

Usage examples:
- Set layout: ppt_layout(operation='set', path='presentation.pptx', slideIndex=0, layout='Title')
- Get layouts: ppt_layout(operation='get_layouts', path='presentation.pptx', masterIndex=0)
- Get masters: ppt_layout(operation='get_masters', path='presentation.pptx')
- Apply master: ppt_layout(operation='apply_master', path='presentation.pptx', slideIndices=[0,1,2], masterIndex=0, layoutIndex=0)
- Apply layout range: ppt_layout(operation='apply_layout_range', path='presentation.pptx', slideIndices=[0,1,2], layout='Title')
- Apply theme: ppt_layout(operation='apply_theme', path='presentation.pptx', themePath='theme.potx')")]
    public string Execute(
        [Description("Operation: set, get_layouts, get_masters, apply_master, apply_layout_range, apply_theme")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for set)")]
        int? slideIndex = null,
        [Description(
            "Layout type: Title, TitleOnly, Blank, TwoColumn, SectionHeader, TitleAndContent, ObjectAndText, PictureAndCaption")]
        string? layout = null,
        [Description("Master index (0-based, optional for get_layouts, required for apply_master)")]
        int? masterIndex = null,
        [Description("Layout index under master (0-based, required for apply_master)")]
        int? layoutIndex = null,
        [Description("Slide indices array as JSON (required for apply_layout_range, optional for apply_master)")]
        string? slideIndices = null,
        [Description("Theme template file path (.potx/.pptx, required for apply_theme)")]
        string? themePath = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, layout, masterIndex, layoutIndex, slideIndices,
            themePath);

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

        var opLower = operation.ToLowerInvariant();
        if (opLower == "get_layouts" || opLower == "get_masters")
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
        string? layout,
        int? masterIndex,
        int? layoutIndex,
        string? slideIndices,
        string? themePath)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLowerInvariant())
        {
            case "set":
                if (slideIndex.HasValue) parameters.Set("slideIndex", slideIndex.Value);
                if (layout != null) parameters.Set("layout", layout);
                break;

            case "get_layouts":
                if (masterIndex.HasValue) parameters.Set("masterIndex", masterIndex.Value);
                break;

            case "get_masters":
                break;

            case "apply_master":
                if (masterIndex.HasValue) parameters.Set("masterIndex", masterIndex.Value);
                if (layoutIndex.HasValue) parameters.Set("layoutIndex", layoutIndex.Value);
                if (slideIndices != null) parameters.Set("slideIndices", slideIndices);
                break;

            case "apply_layout_range":
                if (slideIndices != null) parameters.Set("slideIndices", slideIndices);
                if (layout != null) parameters.Set("layout", layout);
                break;

            case "apply_theme":
                if (themePath != null) parameters.Set("themePath", themePath);
                break;
        }

        return parameters;
    }
}
