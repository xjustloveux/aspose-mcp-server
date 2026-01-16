using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint SmartArt (add, manage nodes)
///     Supports: add, manage_nodes
/// </summary>
[McpServerToolType]
public class PptSmartArtTool
{
    /// <summary>
    ///     Handler registry for SmartArt operations.
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
    ///     Initializes a new instance of the <see cref="PptSmartArtTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptSmartArtTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.SmartArt");
    }

    /// <summary>
    ///     Executes a PowerPoint SmartArt operation (add, manage_nodes).
    /// </summary>
    /// <param name="operation">The operation to perform: add, manage_nodes.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to input path).</param>
    /// <param name="slideIndex">Slide index (0-based, required for all operations).</param>
    /// <param name="shapeIndex">Shape index (0-based, required for manage_nodes).</param>
    /// <param name="layout">
    ///     SmartArt layout type: BasicProcess, BasicCycle, BasicPyramid, BasicRadial, Hierarchy,
    ///     OrganizationChart, etc.
    /// </param>
    /// <param name="x">X position (optional, for add operation, defaults to 100).</param>
    /// <param name="y">Y position (optional, for add operation, defaults to 100).</param>
    /// <param name="width">Width (optional, for add operation, defaults to 400).</param>
    /// <param name="height">Height (optional, for add operation, defaults to 300).</param>
    /// <param name="action">Node action: add, edit, delete (required for manage_nodes operation).</param>
    /// <param name="targetPath">
    ///     Array of indices to target node as JSON (e.g., '[0]' for first node, '[0,1]' for second child
    ///     of first node).
    /// </param>
    /// <param name="text">Node text content (required for add/edit operations in manage_nodes).</param>
    /// <param name="position">Insert position for new node (0-based, optional for add action, defaults to append at end).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_smart_art")]
    [Description(@"Manage PowerPoint SmartArt. Supports 2 operations: add, manage_nodes.

Usage examples:
- Add SmartArt: ppt_smart_art(operation='add', path='presentation.pptx', slideIndex=0, layout='BasicProcess', x=100, y=100, width=400, height=300)
- Manage nodes: ppt_smart_art(operation='manage_nodes', path='presentation.pptx', slideIndex=0, shapeIndex=0, action='add', targetPath='[0]', text='New Node')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a new SmartArt shape (required params: path, slideIndex, layout)
- 'manage_nodes': Manage SmartArt nodes (add, edit, delete) (required params: path, slideIndex, shapeIndex, action)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for all operations)")]
        int slideIndex = 0,
        [Description("Shape index (0-based, required for manage_nodes)")]
        int? shapeIndex = null,
        [Description(
            "SmartArt layout type: BasicProcess, BasicCycle, BasicPyramid, BasicRadial, Hierarchy, OrganizationChart, HorizontalHierarchy, CircleArrowProcess, ClosedChevronProcess, StepDownProcess")]
        string? layout = null,
        [Description("X position (optional, for add operation, defaults to 100)")]
        float x = 100,
        [Description("Y position (optional, for add operation, defaults to 100)")]
        float y = 100,
        [Description("Width (optional, for add operation, defaults to 400)")]
        float width = 400,
        [Description("Height (optional, for add operation, defaults to 300)")]
        float height = 300,
        [Description("Node action: 'add', 'edit', 'delete' (required for manage_nodes operation)")]
        string? action = null,
        [Description(
            "Array of indices to target node as JSON (e.g., '[0]' for first node, '[0,1]' for second child of first node)")]
        string? targetPath = null,
        [Description("Node text content (required for add/edit operations in manage_nodes)")]
        string? text = null,
        [Description("Insert position for new node (0-based, optional for add action, defaults to append at end)")]
        int? position = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, shapeIndex, layout,
            x, y, width, height, action, targetPath, text, position);

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

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int slideIndex,
        int? shapeIndex,
        string? layout,
        float x,
        float y,
        float width,
        float height,
        string? action,
        string? targetPath,
        string? text,
        int? position)
    {
        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(slideIndex, layout, x, y, width, height),
            "manage_nodes" => BuildManageNodesParameters(slideIndex, shapeIndex, action, targetPath, text, position),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the add SmartArt operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="layout">The SmartArt layout type.</param>
    /// <param name="x">X position for the SmartArt.</param>
    /// <param name="y">Y position for the SmartArt.</param>
    /// <param name="width">Width of the SmartArt.</param>
    /// <param name="height">Height of the SmartArt.</param>
    /// <returns>OperationParameters configured for adding SmartArt.</returns>
    private static OperationParameters BuildAddParameters(int slideIndex, string? layout, float x, float y, float width,
        float height)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);
        if (layout != null) parameters.Set("layout", layout);
        parameters.Set("x", x);
        parameters.Set("y", y);
        parameters.Set("width", width);
        parameters.Set("height", height);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the manage_nodes operation.
    /// </summary>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="action">The node action (add, edit, delete).</param>
    /// <param name="targetPath">JSON array of indices to target node.</param>
    /// <param name="text">The node text content.</param>
    /// <param name="position">Insert position for new node.</param>
    /// <returns>OperationParameters configured for managing SmartArt nodes.</returns>
    private static OperationParameters BuildManageNodesParameters(int slideIndex, int? shapeIndex, string? action,
        string? targetPath, string? text, int? position)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (action != null) parameters.Set("action", action);
        if (targetPath != null) parameters.Set("targetPath", targetPath);
        if (text != null) parameters.Set("text", text);
        if (position.HasValue) parameters.Set("position", position.Value);
        return parameters;
    }
}
